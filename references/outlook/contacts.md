# Outlook Contacts

Access contact information and address book features.

## Get Sender Contact (Read Mode)

```javascript
// From property contains sender info
return {
  email: item.from?.emailAddress,
  displayName: item.from?.displayName
};
```

## Get Recipients

### To Recipients

```javascript
// Read mode - direct access
if (item.to) {
  return item.to.map(r => ({
    email: r.emailAddress,
    name: r.displayName
  }));
}

// Compose mode - async
return new Promise(resolve => {
  item.to.getAsync(result => {
    resolve(result.value.map(r => ({
      email: r.emailAddress,
      name: r.displayName
    })));
  });
});
```

### CC Recipients

```javascript
// Read mode
if (item.cc) {
  return item.cc.map(r => ({
    email: r.emailAddress,
    name: r.displayName
  }));
}

// Compose mode
return new Promise(resolve => {
  item.cc.getAsync(result => {
    resolve(result.value.map(r => ({
      email: r.emailAddress,
      name: r.displayName
    })));
  });
});
```

### BCC Recipients (Compose Only)

```javascript
return new Promise(resolve => {
  item.bcc.getAsync(result => {
    resolve(result.value.map(r => ({
      email: r.emailAddress,
      name: r.displayName
    })));
  });
});
```

## Add Recipients

### Add To Recipients

```javascript
return new Promise(resolve => {
  item.to.addAsync([
    { emailAddress: "alice@example.com", displayName: "Alice Smith" },
    { emailAddress: "bob@example.com", displayName: "Bob Jones" }
  ], result => {
    resolve(result.status === Office.AsyncResultStatus.Succeeded);
  });
});
```

### Add CC Recipients

```javascript
return new Promise(resolve => {
  item.cc.addAsync([
    { emailAddress: "manager@example.com", displayName: "Manager" }
  ], result => {
    resolve(result.status === Office.AsyncResultStatus.Succeeded);
  });
});
```

### Add BCC Recipients

```javascript
return new Promise(resolve => {
  item.bcc.addAsync([
    { emailAddress: "archive@example.com", displayName: "Archive" }
  ], result => {
    resolve(result.status === Office.AsyncResultStatus.Succeeded);
  });
});
```

## Set Recipients (Replace All)

```javascript
return new Promise(resolve => {
  item.to.setAsync([
    { emailAddress: "newrecipient@example.com", displayName: "New Recipient" }
  ], result => {
    resolve(result.status === Office.AsyncResultStatus.Succeeded);
  });
});
```

## Get User Profile

```javascript
return {
  email: mailbox.userProfile.emailAddress,
  displayName: mailbox.userProfile.displayName,
  timeZone: mailbox.userProfile.timeZone
};
```

## Resolve Email Address (REST)

```javascript
// Use REST API to resolve/lookup contacts
const token = await new Promise(resolve => {
  Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, result => {
    resolve(result.value);
  });
});

// People API endpoint
const searchEndpoint = `${Office.context.mailbox.restUrl}/v2.0/me/people`;

return {
  endpoint: searchEndpoint,
  query: "$search=Alice&$top=10",
  headers: { Authorization: `Bearer ${token}` }
};
```

## Get Contact Photo (REST)

```javascript
const email = "alice@example.com";
const photoEndpoint = `${Office.context.mailbox.restUrl}/v2.0/users/${email}/photo/$value`;

return {
  endpoint: photoEndpoint,
  responseType: "blob"  // Returns image data
};
```

## Search Contacts (REST)

```javascript
// Search in contacts folder
const contactsEndpoint = `${Office.context.mailbox.restUrl}/v2.0/me/contacts`;
const filter = "$filter=startswith(displayName,'A')";

return {
  endpoint: `${contactsEndpoint}?${filter}`,
  method: "GET"
};
```

## Create Contact (REST)

```javascript
const createEndpoint = `${Office.context.mailbox.restUrl}/v2.0/me/contacts`;

return {
  endpoint: createEndpoint,
  method: "POST",
  body: {
    givenName: "John",
    surname: "Doe",
    emailAddresses: [
      { address: "john.doe@example.com", name: "John Doe" }
    ],
    businessPhones: ["+1-555-123-4567"],
    companyName: "Acme Corp"
  }
};
```

## Update Contact (REST)

```javascript
const contactId = "ABC123...";
const updateEndpoint = `${Office.context.mailbox.restUrl}/v2.0/me/contacts/${contactId}`;

return {
  endpoint: updateEndpoint,
  method: "PATCH",
  body: {
    jobTitle: "Senior Developer"
  }
};
```

## Get Contact Groups (REST)

```javascript
const groupsEndpoint = `${Office.context.mailbox.restUrl}/v2.0/me/contactFolders`;

return {
  endpoint: groupsEndpoint,
  method: "GET"
};
```

## Tips

- Recipients use `{ emailAddress, displayName }` format
- Read mode: direct property access
- Compose mode: use `.getAsync()`, `.setAsync()`, `.addAsync()`
- Contact CRUD operations require REST API
- Use `getCallbackTokenAsync` with `isRest: true` for REST calls
- People API searches across contacts, directory, and recent communications
- BCC is only available in compose mode
