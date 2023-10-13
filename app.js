const express = require('express');
const session = require('express-session');
const msal = require('@azure/msal-node');
const graph = require('@microsoft/microsoft-graph-client');
const {ClientSecretCredential, DefaultAzureCredential} = require('@azure/identity');
const authProviders =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

require('dotenv').config()

const app = express();

app.use(express.json());

const port = process.env.PORT || 3000;

const msalConfig = {
  auth: {
    clientId: 'f662b4fe-6b84-4eed-a986-88597603e865',
    authority: 'https://login.microsoftonline.com/common/',
    clientSecret: '1Zk8Q~DP3f-Z3aZRoDAXNaHm.OEtfsQIFPToqbum',
  },
  system: {
    loggerOptions: {
      loggerCallback(logLevel, message, containsPii) {
        console.log(message);
      },
      piiLoggingEnabled: false,
      logLevel: 'info',
    },
  },
};

const pca = new msal.ConfidentialClientApplication(msalConfig)

const authCodeUrlParameters = {
  scopes: ['User.ReadWrite.All', 'User.Read.All'],
  redirectUri: 'http://localhost:3000/auth/callback',
};

app.use(session({
  secret: '1Zk8Q~DP3f-Z3aZRoDAXNaHm.OEtfsQIFPToqbum',
  resave: false,
  saveUninitialized: true,
}));

const clientId = 'f662b4fe-6b84-4eed-a986-88597603e865';
const clientSecret = '1Zk8Q~DP3f-Z3aZRoDAXNaHm.OEtfsQIFPToqbum';
const tenantId = 'd1ec7c04-0b76-4de7-809e-fba9fcb6e24b';
let client = null;

const tokenCredential = new ClientSecretCredential(tenantId, clientId, clientSecret);

// Create a client with app-only authentication
// app.get('/login', async (req, res) => {
//   if (!client) {
//     const authProvider = new authProviders.TokenCredentialAuthenticationProvider(
//       tokenCredential, {
//         scopes: ['https://graph.microsoft.com/.default'],
//       });
//
//     client = graph.Client.initWithMiddleware({
//       authProvider: authProvider,
//     });
//
//     const response = await tokenCredential.getToken([
//       'https://graph.microsoft.com/.default',
//     ]);
//
//     req.session.authenticated = true;
//     req.session.accessToken = 'eyJhbGciOiJSUzI1NiIsImtpZCI6IjdIOTlsX2xXZTJFRHdhWF90ejBONHBzaXZ4SSIsInR5cCI6IkpXVCIsIng1dCI6IjdIOTlsX2xXZTJFRHdhWF90ejBONHBzaXZ4SSJ9.eyJzZXJ2aWNldXJsIjoiaHR0cHM6Ly9zbWJhLnRyYWZmaWNtYW5hZ2VyLm5ldC9lbWVhLyIsIm5iZiI6MTY5NzEzMjY5MCwiZXhwIjoxNjk3MTM2MjkwLCJpc3MiOiJodHRwczovL2FwaS5ib3RmcmFtZXdvcmsuY29tIiwiYXVkIjoiZDVjNDgzMzktNTA5Yi00NGE0LWI3NzctNjllMGNiNWFiN2ExIn0.HGpdnxtvZmg0ytur1TkWKpaKfC75F1j-Ez87HOfVVUgY5_NmVUCOI3z0Q0LVTuFB8x4KRgSvtWwe_dU8K0z8IFwpHfyTcayzbk8viN0QdLwkT-7Z-RrXxHM53xRAF_z_r6ONro3As4AkPV8gquU5W1I3jt1kAbkULjsOwcOkX6AOZoBUQcl_W6PJRcivKK4xNUQbx854y3538x2z2Zy_2R2xw8RjF0-VzAvwp_oXvdQ50Hfx4XvcDaKxWwWBC2vJWfcomNhuP8HQ2DGovAdvvyzBruJoPbgxXO9wE2dBo4L3d3kpBbvznkz7zV3bvd6tUnQjgYVVrcDG-fhiT61SVg';
//     res.redirect('/');
//   }
// });

app.get('/', (req, res) => {
  res.send(req.session.authenticated ? `${req.session.accessToken}` : 'Logged out');
});

app.get('/login', (req, res) => {
  pca.getAuthCodeUrl(authCodeUrlParameters).then((response) => {
    res.redirect(response);
  }).catch((error) => {
    console.log(error);
    res.status(500).send('Error getting auth code URL');
  });
});

app.get('/auth/callback', (req, res) => {
  const tokenRequest = {
    code: req.query.code,
    scopes: ['User.ReadWrite.All', 'User.Read.All'],
    redirectUri: 'http://0.0.0.0:3000/auth/callback',
  };

  pca.acquireTokenByCode(tokenRequest).then((response) => {
    req.session.authenticated = true;
    req.session.accessToken = response.accessToken;
    res.redirect('/');
  }).catch((error) => {
    console.log(error);
    res.status(500).send('Error acquiring token by code');
  });
});

app.get('/logout', (req, res) => {
  req.session.destroy((err) => {
    res.redirect('/');
  });
});

app.get('/me', async (req, res) => {
  try {
    const accessToken = req.headers.authorization;
    // const response = await tokenCredential.getToken([
    //   'https://graph.microsoft.com/.default',
    // ]);
    // const accessToken = response.token;

    const client =
      graph.Client.init({
        authProvider: (done) => {
          done(null, accessToken);
        },
      });

    const user = await client.api('/me').get();
    console.log(user);
    res.json(user);
  } catch (error) {
    console.error(error);
    res.status(500).send('Error fetching user information');
  }
});

app.get('/users', async (req, res) => {
  try {
    const accessToken = req.headers.authorization;

    const client = graph.Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      },
    });

    const users = await client.api('/users').get();
    console.log(users.value);
    res.json(users.value);
  } catch (error) {
    console.error(error);
    res.status(500).send('Error fetching user list');
  }
});

app.get('/rallies', async (req, res) => {
  try {
    const accessToken = req.headers.authorization;

    const client = graph.Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      },
    });

    // Calculate the start and end dates for the current week
    const today = new Date();
    const startOfWeek = new Date(today);
    startOfWeek.setDate(today.getDate() - today.getDay()); // Start of the week (Sunday)
    startOfWeek.setHours(0, 0, 0, 0);

    const endOfWeek = new Date(today);
    endOfWeek.setDate(today.getDate() + (6 - today.getDay())); // End of the week (Saturday)
    endOfWeek.setHours(23, 59, 59, 999);

    const events = await client.api('/me/calendar/events').get();

    const ralliesThisWeek = events.value.filter((event) => {
      const eventStart = new Date(event.start.dateTime);
      return eventStart >= startOfWeek && eventStart <= endOfWeek;
    });

    res.json(ralliesThisWeek);
  } catch (error) {
    console.error(error);
    res.status(500).send('Error fetching rally events for the current week');
  }
});

app.get('/free-slots/:userId', async (req, res) => {
  try {
    const accessToken = req.headers.authorization;
    const targetUserId = req.params.userId; // The user ID of the target user

    const client = graph.Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      },
    });

    // Calculate the start and end times for the desired time window
    const now = new Date();
    const startTime = new Date(now);
    const endTime = new Date(now);
    endTime.setDate(now.getDate() + 7); // Retrieve free slots for the next 7 days

    // Format start and end times as ISO strings
    const isoStartTime = startTime.toISOString();
    const isoEndTime = endTime.toISOString();

    // Define the query to retrieve free time slots for the target user
    const queryOptions = {
      startDateTime: isoStartTime,
      endDateTime: isoEndTime,
      schedules: [targetUserId], // The ID of the target user
    };

    const freeSlots = await client.api('/me/findMeetingTimes')
      .version('beta')
      .post({...queryOptions});

    res.json(freeSlots);
  } catch (error) {
    console.error(error);
    res.status(500).send('Error fetching free slots');
  }
});

// Define a route to find available meeting times for other users
app.post('/find-meeting-times', async (req, res) => {
  try {
    const accessToken = req.headers.authorization;
    const { startDateTime, endDateTime, userEmails } = req.body;

    const client = graph.Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      },
    });

    // Define the meeting time request
    const meetingTimeRequest = {
      attendees: userEmails.map((email) => ({emailAddress: {address: email}})),
      timeConstraint: {
        activityDomain: 'unrestricted',
        count: 30,
        timeslots: [
          {
            start: {
              dateTime: startDateTime,
              timeZone: 'UTC',
            },
            end: {
              dateTime: endDateTime,
              timeZone: 'UTC',
            },
          },
        ],
      },
      locationConstraint: {
        isRequired: false,
        suggestLocation: false,
      },
      minimumAttendeePercentage: 60,
      meetingDuration: 'PT1H', // 1 hour meeting duration
      returnSuggestionReasons: true,
    };

    // Find available meeting times
    const meetingTimes = await client.api('/me/findMeetingTimes').post(meetingTimeRequest);

    res.json(meetingTimes.meetingTimeSuggestions);
  } catch (error) {
    console.error(error);
    res.status(500).send('Error finding meeting times');
  }
});

app.post('/edit-meeting', async (req, res) => {
  try {
    const accessToken = req.headers.authorization;
    const { startDateTime, endDateTime, meetingId } = req.body;

    const client = graph.Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      },
    });

    // Define the event ID to be edited
    const eventId = req.params.eventId;

    // Define the updated meeting details
    const updatedMeetingDetails = {
      start: {
        dateTime: startDateTime,
        timeZone: 'UTC',
      },
      end: {
        dateTime: endDateTime,
        timeZone: 'UTC',
      },
    };

    // Update the meeting using the PATCH method
    await client.api(
      `/me/events/${meetingId}`)
      .patch(updatedMeetingDetails);

    res.send('Meeting updated successfully');
  } catch (error) {
    console.error(error);
    res.status(500).send('Error updating meeting');
  }
});

app.post('/create-meeting', async (req, res) => {
  try {
    const accessToken = req.headers.authorization;

    const client = graph.Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      },
    });
    const {participants} = req.body;

    const attendees = [];
    participants.forEach(participant => {
      attendees.push({emailAddress: {address: participant}});
    });

    // Define the meeting details for the new meeting
    const newMeetingDetails = {
      subject: req.body.subject,
      start: {
        dateTime: req.body.startDateTime,
        timeZone: 'UTC',
      },
      end: {
        dateTime: req.body.endDateTime,
        timeZone: 'UTC',
      },
      attendees,
    };

    // Create the new meeting using the POST method
    const createdMeeting = await client.api('/me/events').post(newMeetingDetails);

    res.json(createdMeeting);
  } catch (error) {
    console.error(error);
    res.status(500).send('Error creating meeting');
  }
});

// Middleware to ensure authentication before accessing the /me endpoint
function ensureAuthenticated(req, res, next) {
  if (req.isAuthenticated()) {
    return next();
  }
  res.redirect('/login');
}

app.listen(process.env.PORT, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
