# Pylighter
One of the biggest parts of my current job is printing out our events list every single Saturday and highlighting lines according to their information (event/meeting name, location, concierge duties, etc.).
When we get larger lists, this can be incredibly time consuming. After looking into Python and Python-docx, I realized we could use it to shorten this process.
The result? Pylighter. 

Our standard highlighting goes: <br />
PINK - Day and Date <br />
ORANGE - Event/Meeting Title <br />
YELLOW - Event Date/Time <br />
GREEN - specified Concierge duties <br />

However, the one major flaw is that Word does not offer an orange highlighter, so red is what is substituted.

A Brief How-To:
1. Ensure that the text is in the correct format. We use this: <br />
 *(Events are sorted by days, with each section starting with the underlined day and date)* 
  
    **Event:** event title here <br />
    **Event Date/Time:** date of event <br />
    **Location:** room <br />
    **Organization:** either within company or outside org.<br />
    -other details pertaining to event-
    **Concierge / other departments:** what we/they will do <br />

    If copied & pasted, they must be pasted with the "Keep Text Only" setting, as there are
    weird things that can occur with the standard Ctrl + V pasting, involving new lines and
    spacing, that can't be worked around very easily. <br />
3. When saving, make sure you save it as a docx file. Otherwise, this will not work. It's highly recommended to go over the inputted information to make sure there are no typos, missing bits, or anything else important. <br />
4. Start the program. It will ask you what file you'd wish to open. Select where to save and you're done! <br />
