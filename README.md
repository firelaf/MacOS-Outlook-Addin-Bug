## MacOS Desktop Outlook Add-In Dialog Bug

This is a demonstration of a bug with the MacOS version of Outlook. When creating a dialog box from an Add-in, that diaglog box cannot access the Office.RoamingSettings API. This behavior has been found to function on both the Windows Desktop and Web versions of Outlook. However, I'm not sure whether being able to access Roaming Settings from a dialog is intended because it isn't documented anywhere in Microsoft's documentation.

#### Running the demo

Simply run `npm install`, and then `npm run build`.

To start the application, sideload automatically by running `npm start` or alternatively, sideload manually by running `npm run dev-server` and add the Add-in from file, chosing the manifest.xml inside the root folder of the repo.

Once sideloaded, press the Open Dialog button inside of Outlook. Follow the instructions on the dialog screen - attempt to set the boolean value in Roaming Settings, then close the dialog and either open it back up, or open the taskpane to see if the value has been retained. This shows how the dialog cannot set roaming settings. If you're on Windows or Web, it will be. If you're on MacOS Desktopk, it will not.

Another thing to try is to opent the taskpane, change the value, and then open the dialog and see if the the value has been changed. This shows how the dialog cannot read roaming settings either.
