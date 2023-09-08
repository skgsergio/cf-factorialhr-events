import ICAL from 'ical.js';

const team = [
  "Jose Luis Domínguez Balirac",
  "David Garcinuño Enríquez",
  "Samuel García Haro",
  "Sergio Valverde García",
  "Sergio Conde Gómez"
];

function datePad(n) {
  return ('0' + n).slice(-2);
}

function dayFmt(d) {
  return datePad(d.getDate()) + '-' +
    datePad(d.getMonth() + 1) + '-' +
    d.getFullYear();
}

function dateFmt(d) {
  const tzDiff = -d.getTimezoneOffset();

  return dayFmt(d) + ' ' +
    datePad(d.getHours()) + ':' +
    datePad(d.getMinutes()) + ':' +
    datePad(d.getSeconds()) + ' ' +
    (tzDiff >= 0 ? '+' : '-') +
    datePad(Math.floor(Math.abs(tzDiff) / 60)) + ':' +
    datePad(Math.abs(tzDiff) % 60);
}

async function getEvents(token) {
  const todayTeamEvents = [];
  const currentDate = new Date();

  const result = await fetch(`https://factorialhr.com/icals?token=${token}`);
  const icalData = await result.text();

  const comp = new ICAL.Component(ICAL.parse(icalData));
  const events = comp.getAllSubcomponents('vevent');
  const tz = new ICAL.Timezone(comp.getFirstSubcomponent("vtimezone"));

  events.forEach(eventPlain => {
    const event = new ICAL.Event(eventPlain);

    event.startDate.zone = tz;
    const startDate = event.startDate.toJSDate();

    event.endDate.zone = tz;
    const endDate = event.endDate.toJSDate();

    if (
      team.some(member => event.summary.includes(member)) &&
      !event.summary.includes("aniversario") &&
      startDate <= currentDate &&
      endDate > currentDate
    ) {
      if ((endDate - startDate) == 86400000) { // Single day
        todayTeamEvents.push(`${dayFmt(startDate)} :: ${event.summary}`);
      } else if ((endDate - startDate) % 86400000 == 0) { // Multiple complete days
        startDate.setMinutes(startDate.getMinutes() + startDate.getTimezoneOffset());
        endDate.setMinutes(endDate.getMinutes() + endDate.getTimezoneOffset());
        endDate.setSeconds(endDate.getSeconds() - 1);
        todayTeamEvents.push(`${dayFmt(startDate)} - ${dayFmt(endDate)} :: ${event.summary}`);
      } else { // Non complete day events
        todayTeamEvents.push(`${dateFmt(startDate)} - ${dateFmt(endDate)} :: ${event.summary}`);
      }
    }
  });

  return todayTeamEvents;
}

async function sendTeamsMessage(whUrl, events) {
  if (events.length == 0) {
    return;
  }

  const result = await fetch(
    whUrl,
    {
      method: "POST",
      body: JSON.stringify({
        "type": "message",
        "attachments": [
          {
            "contentType": "application/vnd.microsoft.card.adaptive",
            "contentUrl": null,
            "content": {
              "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
              "type": "AdaptiveCard",
              "version": "1.5",
              "body": [
                {
                  "type": "TextBlock",
                  "text": (new Date()).toISOString().split('T')[0],
                  "wrap": true,
                  "style": "heading"
                }
              ].concat(events.map(event => ({
                "type": "TextBlock",
                "wrap": true,
                "text": event
              })))
            }
          }
        ]
      })
    }
  );

  if (!result.ok) {
    throw new Error("Teams request failed: " + (await result.text()));
  }
}

export default {
  async fetch(request, env, context) {
    const evts = await getEvents(env.FHR_TOKEN);
    await sendTeamsMessage(env.TEAMS_WH, evts)
    return new Response("OK - Check Teams");
  },
  async scheduled(event, env, ctx) {
    const evts = await getEvents(env.FHR_TOKEN);
    await sendTeamsMessage(env.TEAMS_WH, evts)
  },
};
