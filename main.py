import pyautogui as bot
import win32clipboard as cb


# Initial settings configuration
screen = bot.size()
DEFAULT_PAUSE = 1.5
bot.FAILSAFE = True
bot.PAUSE = DEFAULT_PAUSE

# Basic details, always update the course and section
course_section = bot.prompt("Enter the full course and  section:\nEx. BSIT 2D")
full_name = ""
semester = ""
REQUIRED_GWA = 1.75
is_one_line = False

# List of screen coordinates
CHROME_1ST_TAB = (123, 16)
CHROME_2ND_TAB = (338, 3)
NAME_BOX = (713, 272)
UPPER_1ST_SEM_LAST = (837, 335)
UPPER_2ND_SEM_LAST = (856, 338)
LOWER_1ST_SEM_LAST = (837, 367)
LOWER_2ND_SEM_LAST = (856, 367)
GWA_BOX = (792, 462)
DOWNLOAD_BUTTON = (1170, 125)
DOWNLOAD_CONFIRM_BUTTON = (1202, 466)
CLOSE_ADS = (1128, 162)
CLOSE_DOWNLOADS_TRAY = (1344, 695)
DOWNLOAD_UP_SYMBOL = (210, 701)
SHOW_IN_FOLDER = (251, 634)
CLOSE_EXPLORER = (1047, 171)
SCROLL_DOWN_CLICK = (1354, 655)


def get_clipboard():
    cb.OpenClipboard()
    data = cb.GetClipboardData()
    cb.CloseClipboard()
    return data


def set_clipboard(data):
    cb.OpenClipboard()
    cb.EmptyClipboard()
    cb.SetClipboardText(data)
    cb.CloseClipboard()


def convert_name(name):
    # Converting name if it doesn't fit the format (First Name, MI, Last Name)
    comma = name.find(',')
    if comma == -1: return name
    return (name[comma + 1:] + " " + name[0:comma]).strip().replace(',','')


def center_cursor():
    # Usage is mostly for resetting the controls or to invoke the bot.PAUSE
    bot.moveTo(screen[0] / 2 , screen[1] / 2)


def remove_semester():
    bot.PAUSE = 0.1
    # Just a backspace function to delete the semester's info on the screen
    for x in range(3): bot.press('backspace')
    bot.PAUSE = DEFAULT_PAUSE


def save_template(semester):
    # Clicking the download button in the upper left of the website
    bot.click(x=DOWNLOAD_BUTTON[0], y=DOWNLOAD_BUTTON[1])

    # Changing the pause duration because the screen will load for a bit
    bot.PAUSE = 9

    # Confirming the download
    bot.click(x=DOWNLOAD_CONFIRM_BUTTON[0], y=DOWNLOAD_CONFIRM_BUTTON[1])

    # Reverting it back
    bot.PAUSE = DEFAULT_PAUSE

    # Clicking the download span button to view it on explorer later
    bot.click(x=DOWNLOAD_UP_SYMBOL[0], y=DOWNLOAD_UP_SYMBOL[1])
    bot.click(x=SHOW_IN_FOLDER[0], y=SHOW_IN_FOLDER[1])

    # Imitating the right click by doing the shortcut of it
    bot.hotkey('shift', 'f10')

    # Hotkey for 'rename'
    bot.press('m')

    # Setting the file name and saving it on clipboard
    set_clipboard(f"{full_name} - ({course_section} - {semester} Sem)")

    # Pasting the name
    bot.hotkey('ctrl', 'v')

    # Saving the name
    bot.press('enter')

    # Close the downloads window
    bot.click(x=CLOSE_EXPLORER[0], y=CLOSE_EXPLORER[1])

    # Closing the downloads tray first
    bot.click(x=CLOSE_DOWNLOADS_TRAY[0], y=CLOSE_DOWNLOADS_TRAY[1])

    # Then closing the ads
    bot.click(x=CLOSE_ADS[0], y=CLOSE_ADS[1])

    # Going back to the spreadsheet, end of the function
    bot.click(x=CHROME_1ST_TAB[0], y=CHROME_1ST_TAB[1])



if __name__ == "__main__":
    # A confirmation window will show up before starting the script
    if bot.confirm("Ready?") != "OK": quit()
    center_cursor()

    #            The loop will start here
    # The screen must be on the 1st chrome window, selected the first name
    # on the excel spreadsheet to begin, and the semester temporarily must be
    # on the 2nd semester initially, in order for smooth looping
    
    while True:
        # Copy the whole name
        bot.hotkey('ctrl','c')

        # Test if the cell is empty, the size 2 is reasonable because there will be
        # name that consists only of 2 letters
        if len(get_clipboard()) == 2:
            bot.alert("The process stopped because it reached an empty cell.")
            quit()

        # Reformat name so it will follow the (First Name, MI, Last Name) format required
        full_name = convert_name(get_clipboard())
        
        # Updating the clipboard to the updated name
        set_clipboard(full_name)

        # Going to the second tab
        bot.click(x=CHROME_2ND_TAB[0], y=CHROME_2ND_TAB[1])

        # Scrolling down to the lowest
        bot.doubleClick(x=SCROLL_DOWN_CLICK[0], y=SCROLL_DOWN_CLICK[1])

        # Double clicking on the name box
        bot.doubleClick(x=NAME_BOX[0], y=NAME_BOX[1])

        # Select all / Highlight
        bot.hotkey('ctrl', 'a')

        # Paste from clipboard
        bot.hotkey('ctrl', 'v')

        # Test: Determining the location of the 1st/2nd on the screen.
        # First, moving to the assumed location if the name only spans one line
        bot.doubleClick(x=UPPER_2ND_SEM_LAST[0], y=UPPER_2ND_SEM_LAST[1])

        # Selecting all, to make sure it is all selected for comparison purposes.
        bot.hotkey('ctrl','a')

        # Copying so that the value will be stored in the clipboard before comparing
        bot.hotkey('ctrl', 'c')

        # Doing the test, comparing it to the variable full_name
        if get_clipboard() == full_name:
            # It means that the name spans 2 lines
            is_one_line = False
            # It is the same, therefore using the lower coordinates
            bot.doubleClick(x=LOWER_2ND_SEM_LAST[0], y=LOWER_2ND_SEM_LAST[1])
        else:
            # The name only span one line
            is_one_line = True
            # It isn't the same, so the system will use the upper coordinates
            bot.click(x=UPPER_2ND_SEM_LAST[0], y=UPPER_2ND_SEM_LAST[1])

        # Typing the updated value
        remove_semester()
        bot.typewrite("1st", interval=0.1)

        # Now going back to the excel in order to get the value of the student's GWA
        bot.click(x=CHROME_1ST_TAB[0], y=CHROME_1ST_TAB[1])

        # Moving a cell one to the right, to focus on the first semester's GWA
        bot.press('right')

        # Copying the GWA
        bot.hotkey('ctrl', 'c')

        # Testing if the GWA passes the required GWA (1.75)
        if float(get_clipboard()) <= REQUIRED_GWA:
            # Going back to the template
            bot.click(x=CHROME_2ND_TAB[0], y=CHROME_2ND_TAB[1])

            # Traversing to the GPA box, double click twice to highlight the value
            bot.doubleClick(x=GWA_BOX[0], y=GWA_BOX[1])
            bot.doubleClick(x=GWA_BOX[0], y=GWA_BOX[1])

            # Pasting the First Semester's GWA
            bot.hotkey('ctrl', 'v')

            save_template('First')

        # Moving a cell one to the right, to focus on the second semester's GWA
        bot.press('right')

        # Copying the GWA
        bot.hotkey('ctrl', 'c')

        # Testing if the GWA passes the required GWA (1.75)
        if float(get_clipboard()) <= REQUIRED_GWA:
            # Going back to the template
            bot.click(x=CHROME_2ND_TAB[0], y=CHROME_2ND_TAB[1])

            # Scrolling down to the lowest, double checking
            bot.doubleClick(x=SCROLL_DOWN_CLICK[0], y=SCROLL_DOWN_CLICK[1])

            # Getting the previous settings of coordinates based on the lines accumulated
            # by the whole name of the student
            if is_one_line:
                # Same explanation as of before
                bot.doubleClick(x=UPPER_1ST_SEM_LAST[0], y=UPPER_1ST_SEM_LAST[1])
            else:
                bot.doubleClick(x=LOWER_1ST_SEM_LAST[0], y=LOWER_1ST_SEM_LAST[1])

            # Typing the updated value
            remove_semester()
            bot.typewrite("2nd", interval=0.1)

            # Traversing to the GPA box, double click twice to highlight the value
            bot.doubleClick(x=GWA_BOX[0], y=GWA_BOX[1])
            bot.doubleClick(x=GWA_BOX[0], y=GWA_BOX[1])

            # Pasting the First Semester's GWA
            bot.hotkey('ctrl', 'v')

            save_template('Second')
    
        # Moving the selected cell to the next student, setting everything from
        # the start, to iterate over the list
        bot.press('down')
        bot.press('left')
        bot.press('left')
    # The loop ends here
