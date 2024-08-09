# cwNotify
 ConnectWise Notifier

# Setup
On first launch you'll be asked for your information.
- For API Keys login to ConnectWise, Goto my Account from top right hand menu, click API Keys, Click +, enter cwNotify and then save (not save+Close). Make note of the pub/priv keys only for this setup. WARN: AFter setup is complete discard the API keys, do not save them.
- Get the ClientId from me via Teams.
- Your username is usually your email alias without @domain.tld, but internal to ConnectWise it may be different.
- cwNotify will validate the API keys when you click on and then start watching tickets.
- cwNotify encrypts these details and they can only be decrypted on the profile/machine they were encrypted on. (CryptProtectData API)
