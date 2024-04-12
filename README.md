This is a trial project of key management and tracking based out of google sheets. May end up moving to a different program, but the framwork is already setup

Features
- Add Keys to assignment or to a Lockbox
- Remove keys from assignment to a Lockbox
- Safety features to ensure no unwanted transfer or movement of keys that is not documented
- Define a restricted key and create rules on it's logic
  - (/^[A-F]\d{1,2}$/i test.name)
  - Can never have more than one instance in existance
 


Upcoming
- Define Non-restricted keys and set rules
- Set definition for lost keys
- require name when filling out a form


Current
- Non restricted keys arent getting the formObject information properly. So while the restricted key form is working properly (Which shows that processform function is working) the non restricted key form function is not working and may need to be rewritten.
Last change was just troubleshooting where the info is getting lost, and how it's getting transfered currently
