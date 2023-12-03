# Pylighter
One of the biggest parts of my current job is printing out our events list every single Saturday and highlighting lines according to their information (event/meeting name, location, concierge duties, etc.).
When we get larger lists, this can be incredibly time consuming. After looking into Python and Python-docx, I realized we could use it to shorten this process.
The result? Pylighter. 

Our standard highlighting goes:
PINK - Day and Date;
ORANGE - Event/Meeting Title;
YELLOW - Event Date/Time;
GREEN - specified Concierge duties

However, the one major flaw is that Word does not offer an orange highlighter, so red is what is substituted.
