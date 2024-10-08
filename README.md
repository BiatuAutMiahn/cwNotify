# cwNotify
 ConnectWise Notifier

# Setup
On first launch you'll be asked for your information.
- For API Keys login to ConnectWise, Goto my Account from top right hand menu, click API Keys, Click +, enter cwNotify and then save (not save+Close). Make note of the pub/priv keys only for this setup. WARN: After setup is complete discard the API keys, do not save them.
- Get the ClientId from me via Teams.
- Your username is usually your email alias without @domain.tld, but internal to ConnectWise it may be different.
- cwNotify will validate the API keys and then start watching tickets.
- cwNotify encrypts the API details and they can only be decrypted on the user profile and device they were encrypted on. (CryptProtectData API) So in the event that you use the notifier on a different device you will need to generate new API keys.

# Function
cwNotfy works syncronously by retrieving a list of ticket ids /w the last update time stamp for service/project tickets that you have been assigned to or are a resource of, and the same for the tickets that have been recieved previously. Next it will sequentially gather full details about tickets that have a different time stamp from the last ones. For each ticket that has been updated, it will generate text outlining the differences and then show a notification dialog. 

# Caveats
- The notification dialog blocks the code flow of the script until you acknowledge it. ie; cwNotify cannot look for updates while the notification is active.
- Intermediate changes between update checks are not tracked. ie; If a ticket gets updated 4 times since the last notification, you'll only be notified of the last update when cwNotify does it's next check.
- The last note shown is the note at the top of the ticket notes, not the last note entered chronologically.
- The loop interval is 10 sec, getting the list of ticket ids takes anywhere from 10 to 20 seconds, so you can expect ticket updates within a minute and soemtimes within 10 seconds.
- Once a ticket is in one of the closed statuses and is older than 7 days it will be removed from the watch list.
- Some ticket updates will update the last modified time stamp and the notification will not show the field that changed. ie; resource acknowledged, resource added/removed/marked done etc. cwNotify only monitors the following fields: `_info.dateEntered` `_info.lastUpdated` `id` `status.name` `owner.name` `summary` `company.name` `contact.name` `subType.name` `item.name` `priority.name` `severity.name` `type.name` `_info.updatedBy`

# TODO
- Async ticket updates (libcurl)
- Queue update notifications (for non-blocking notify)
- Summary dialog
- Snooze Button
- Suspend updates option
- Auto dismiss option
- Manually add a ticket to watch
- option to ignore/stop monitoring a ticket.
- Rework notify box /w colored text and proper DPI scaling.

# Release Candidate
`_.rc\cwNotify.async.exe` in the newFetch branch.
- Notifications are non-blocking.
- Ticket retrieval is asynchronous.
