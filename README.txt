Job Packet Tracker

Ah, VB6. So out of date, yet so fun and familiar.

Job Packet Tracker is an application designed to improve office productivity by keeping track of Job Packet movements between different departments and hands. It works as a simple MySQL front-end.

A job packet is a physical folder that contains all pertinent information for a single job. This includes drawings, quotes, BOMs, customer info, etc.
Job Packet Tracker works with a simple MySQL database via ODBC. It tracks each physical movement of the job packet. The states include Created, Sent, Received, Filed, Reopened and Closed. A user enters the job number to pull up the packet, then selects the desired action and username (if needed) and submits the update. Most commonly, users will be sending to another user who then receives it. It also has a notation feature where users can add short messages for each update that is stored in the DB.

Job Packet Tracker takes this information and stores it in a very simple database.  This information is then used in several ways to provide more useful info to the end-user. Most notably, detailed search functions, current packets in possession, current packets on the way, and a history record which includes a timeline that displays how long the packet has been in each state.
  
Another feature includes what I call “Live Search”, which was inspired by Google Instant Search. It works by waiting X amount of time after the last character was keyed into the job number field then makes a short query containing that info. If any matches are found, they are displayed in a listbox directly below the job number textbox, users are free to keydown or click on any of these items to pull up the corresponding packet. 

Other features include local alerts for new incoming packets, which are displayed via another feature I call “Banner”. The Banner, also sometimes called “Slider” is a user-control-esk feature I wrote which uses a frame that slides out from the top edge of the window to display whatever info I need. It’s customizable by text, color, duration and click actions. It also caches multiple calls, then displays them one at a time as they timeout, or are clicked.

In the end I wanted to reduce the number of clicks to a few as possible.   It’s possible to send a packet to a user in 3 clicks.

To be continued?
