# Outlook Folders

Navigate folders and manage mail organization.

## Get Current Folder

```javascript
// REST API approach for folder access
const restUrl = Office.context.mailbox.restUrl;
const token = await new Promise(resolve => {
  Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, result => {
    resolve(result.value);
  });
});

return { restUrl, hasToken: !!token };
```

## Get User Mailbox Info

```javascript
return {
  email: mailbox.userProfile.emailAddress,
  displayName: mailbox.userProfile.displayName,
  timeZone: mailbox.userProfile.timeZone,
  accountType: mailbox.userProfile.accountType
};
```

## Get Current Item Context

```javascript
// Information about where the item lives
return {
  itemType: item.itemType,
  itemId: item.itemId,
  conversationId: item.conversationId,
  internetMessageId: item.internetMessageId
};
```

## Well-Known Folders

Outlook has standard folder names accessible via REST API or EWS:

- `Inbox` - Incoming mail
- `Drafts` - Draft messages
- `SentItems` - Sent mail
- `DeletedItems` - Trash
- `Archive` - Archived mail
- `JunkEmail` - Spam folder
- `Outbox` - Pending send

## Move Item (via REST)

```javascript
// Get REST token
const token = await new Promise(resolve => {
  Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, result => {
    resolve(result.value);
  });
});

// Move requires REST API call
// This shows the concept - actual implementation needs fetch
const moveEndpoint = `${Office.context.mailbox.restUrl}/v2.0/me/messages/${item.itemId}/move`;

return {
  endpoint: moveEndpoint,
  method: "POST",
  headers: { Authorization: `Bearer ${token}` },
  body: { destinationId: "Archive" }
};
```

## Copy Item

```javascript
// Similar to move, uses REST API
const copyEndpoint = `${Office.context.mailbox.restUrl}/v2.0/me/messages/${item.itemId}/copy`;

return {
  endpoint: copyEndpoint,
  method: "POST",
  body: { destinationId: "Inbox" }
};
```

## Get Folder List (REST)

```javascript
// Build REST endpoint for folder listing
const foldersEndpoint = `${Office.context.mailbox.restUrl}/v2.0/me/mailFolders`;

const token = await new Promise(resolve => {
  Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, result => {
    resolve(result.value);
  });
});

return {
  endpoint: foldersEndpoint,
  token: token.substring(0, 20) + "..."  // Truncated for display
};
```

## Create Folder (REST)

```javascript
// REST API endpoint for folder creation
const createEndpoint = `${Office.context.mailbox.restUrl}/v2.0/me/mailFolders`;

return {
  endpoint: createEndpoint,
  method: "POST",
  body: {
    displayName: "Project Alpha"
  }
};
```

## Get Subfolder (REST)

```javascript
// Get child folders of a folder
const parentFolderId = "Inbox";
const subfolderEndpoint = `${Office.context.mailbox.restUrl}/v2.0/me/mailFolders/${parentFolderId}/childFolders`;

return { endpoint: subfolderEndpoint };
```

## Search in Folder (REST)

```javascript
// Search messages in a folder
const searchEndpoint = `${Office.context.mailbox.restUrl}/v2.0/me/mailFolders/Inbox/messages`;
const filter = "$filter=contains(subject,'Project')";

return {
  endpoint: `${searchEndpoint}?${filter}`,
  method: "GET"
};
```

## Get Unread Count (REST)

```javascript
// Get folder with unread count
const inboxEndpoint = `${Office.context.mailbox.restUrl}/v2.0/me/mailFolders/Inbox`;

return {
  endpoint: inboxEndpoint,
  select: "$select=displayName,unreadItemCount,totalItemCount"
};
```

## Working with Categories

```javascript
// Get categories (read mode)
if (item.categories) {
  return item.categories;  // Array of category names
}

// Categories in compose mode
return new Promise(resolve => {
  item.categories.getAsync(result => {
    resolve(result.value);
  });
});
```

## Add Category

```javascript
return new Promise(resolve => {
  item.categories.addAsync(["Important", "Follow Up"], result => {
    resolve(result.status === Office.AsyncResultStatus.Succeeded);
  });
});
```

## Tips

- Folder operations primarily use REST API, not Office.js
- Use `getCallbackTokenAsync` with `isRest: true` for REST calls
- Well-known folder names are locale-independent
- Categories are applied to items, not folders
- REST API requires appropriate permissions in manifest
- For complex folder operations, consider Microsoft Graph API
