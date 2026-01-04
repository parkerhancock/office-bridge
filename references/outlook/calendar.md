# Outlook Calendar

Work with appointments, meetings, and calendar events.

## Detect Appointment Context

```javascript
// Check if current item is an appointment
return {
  type: item.itemType,  // "appointment" or "message"
  isAppointment: item.itemType === Office.MailboxEnums.ItemType.Appointment
};
```

## Read Mode - Get Appointment Details

```javascript
// Direct property access in read mode
return {
  subject: item.subject,
  start: item.start,
  end: item.end,
  location: item.location,
  organizer: item.organizer?.emailAddress,
  requiredAttendees: item.requiredAttendees?.map(a => a.emailAddress),
  optionalAttendees: item.optionalAttendees?.map(a => a.emailAddress)
};
```

## Compose Mode - Get Appointment Details

```javascript
// Async access in compose mode
const getAsync = (prop) => new Promise(resolve => {
  item[prop].getAsync(result => resolve(result.value));
});

const subject = await getAsync('subject');
const start = await getAsync('start');
const end = await getAsync('end');
const location = await getAsync('location');

return { subject, start, end, location };
```

## Set Appointment Subject

```javascript
return new Promise(resolve => {
  item.subject.setAsync("Team Sync - Weekly", result => {
    resolve(result.status === Office.AsyncResultStatus.Succeeded);
  });
});
```

## Set Appointment Time

```javascript
// Set start time
const startTime = new Date("2024-12-15T10:00:00");
await new Promise(resolve => {
  item.start.setAsync(startTime, result => resolve(result.status));
});

// Set end time
const endTime = new Date("2024-12-15T11:00:00");
await new Promise(resolve => {
  item.end.setAsync(endTime, result => resolve(result.status));
});

return "Times set";
```

## Set Location

```javascript
return new Promise(resolve => {
  item.location.setAsync("Conference Room A", result => {
    resolve(result.status === Office.AsyncResultStatus.Succeeded);
  });
});
```

## Add Required Attendees

```javascript
return new Promise(resolve => {
  item.requiredAttendees.addAsync([
    { emailAddress: "alice@example.com", displayName: "Alice Smith" },
    { emailAddress: "bob@example.com", displayName: "Bob Jones" }
  ], result => {
    resolve(result.status === Office.AsyncResultStatus.Succeeded);
  });
});
```

## Add Optional Attendees

```javascript
return new Promise(resolve => {
  item.optionalAttendees.addAsync([
    { emailAddress: "carol@example.com", displayName: "Carol White" }
  ], result => {
    resolve(result.status === Office.AsyncResultStatus.Succeeded);
  });
});
```

## Get All Attendees

```javascript
const getAttendees = (prop) => new Promise(resolve => {
  item[prop].getAsync(result => {
    resolve(result.value.map(a => ({
      email: a.emailAddress,
      name: a.displayName
    })));
  });
});

const required = await getAttendees('requiredAttendees');
const optional = await getAttendees('optionalAttendees');
return { required, optional };
```

## Set Appointment Body

```javascript
return new Promise(resolve => {
  item.body.setAsync(
    "Agenda:\n1. Project updates\n2. Q&A\n3. Next steps",
    { coercionType: Office.CoercionType.Text },
    result => {
      resolve(result.status === Office.AsyncResultStatus.Succeeded);
    }
  );
});
```

## Set HTML Body

```javascript
return new Promise(resolve => {
  item.body.setAsync(
    "<h2>Meeting Agenda</h2><ul><li>Updates</li><li>Discussion</li></ul>",
    { coercionType: Office.CoercionType.Html },
    result => {
      resolve(result.status === Office.AsyncResultStatus.Succeeded);
    }
  );
});
```

## Get Recurrence (Read Mode)

```javascript
// Only available for recurring appointments
if (item.recurrence) {
  return {
    recurrenceType: item.recurrence.recurrenceType,
    seriesTime: item.recurrence.seriesTime,
    recurrenceProperties: item.recurrence.recurrenceProperties
  };
}
return "Not a recurring appointment";
```

## Set Recurrence (Compose Mode)

```javascript
// Set weekly recurrence
const recurrence = {
  recurrenceType: Office.MailboxEnums.RecurrenceType.Weekly,
  recurrenceProperties: {
    interval: 1,
    days: [Office.MailboxEnums.Days.Monday, Office.MailboxEnums.Days.Wednesday]
  },
  seriesTime: {
    startDate: "2024-01-01",
    endDate: "2024-06-30",
    startTime: "10:00",
    duration: 60  // minutes
  }
};

return new Promise(resolve => {
  item.recurrence.setAsync(recurrence, result => {
    resolve(result.status === Office.AsyncResultStatus.Succeeded);
  });
});
```

## Get Free/Busy Status

```javascript
// In read mode
return item.showAs;  // "free", "tentative", "busy", "oof", "workingElsewhere"
```

## Tips

- Appointments use `item.itemType === "appointment"`
- Read mode: direct property access
- Compose mode: use `.getAsync()` and `.setAsync()`
- Time zones are handled automatically by Office
- Recurrence patterns can be complex - test thoroughly
- Organizer info is read-only
