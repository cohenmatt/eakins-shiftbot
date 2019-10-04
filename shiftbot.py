"""
Checks most recent emails from a Gmail account for Eakins shift offerings,
and takes any shifts that fit the user's date preferences.

Author:     Matt Cohen
Date:       30/05/2018
Version:    1.0
"""

import smtplib
import time
import imaplib
import email
import webbrowser
import re

# Set credentials
ORG_EMAIL   = "@gmail.com"
FROM_EMAIL  = "" + ORG_EMAIL
FROM_PWD    = ""
SMTP_SERVER = "imap.gmail.com"
SMTP_PORT   = 993

# Eakins email details
EAKINS_SUBJECT  = "A shift has been offered up on the Catering roster"
EAKINS_SENDER   = "\"QI on behalf of noreply@queens.unimelb.edu.au\" <www-data@queens.unimelb.edu.au>"

# Shift preferences
PREF_DAYS   = ["Sat", "Sun"]
PREF_MONTHS = ["Jun", "Jul"]
PREF_YEARs  = ["2018"]

def gmail_login():
    """Log in to Gmail account."""

    try:
        # Log in to Gmail
        mail = imaplib.IMAP4_SSL(SMTP_SERVER)
        mail.login(FROM_EMAIL,FROM_PWD)

        # Select the inbox we want, don't mark as read - just view it
        mail.select('inbox', readonly=True)

        return mail

    except Exception, e:
        print str(e)

def date_check(shift_time):
    """Checks to see if a shift is at a valid time. Currently doesn't account for
    the hourly time of the shift."""

    time_list = shift_time.group(1).split()

    # Days
    if (time_list[1][:-1] not in PREF_DAYS):
        return False

    # Months
    if (time_list[3] not in PREF_MONTHS):
        return False

    # Years
    if (time_list[4] not in PREF_YEARs):
        return False

    # All checks successfull
    return True


def extract_body(payload):
    """Extracts the body of a given email and returns it as a string."""
    if isinstance(payload, str):
        return payload
    else:
        return '\n'.join([extract_body(part.get_payload()) for part in payload])


def read_email():
    """Filters through emails from a given address, finds ones notifying the
    recipient about an Eakins shift on offer, and takes the shift if it fits
    the recipient's preferences for shift times."""
    try:
        # Log in to Gmail
        mail = gmail_login()

        # Get a list of IDs for every email on the account
        type, data = mail.search(None, 'ALL')
        mail_ids = data[0]

        # Get the IDs of the first and last emails
        id_list = mail_ids.split()
        first_email_id = int(id_list[0])
        latest_email_id = int(id_list[-1])

        # Search last 20 emails
        for i in range(latest_email_id, latest_email_id - 3, -1):

            # Get this specific email (the first one)
            type, data = mail.fetch(i, '(RFC822)')

            # Print its values
            for response_part in data:
                if isinstance(response_part, tuple):

                    # Get email subject and sender
                    msg = email.message_from_string(response_part[1])
                    email_subject = msg['subject']
                    email_from = msg['from']

                    # Test that we're reading in emails
                    # print "Subject:     " + email_subject
                    # print "From:        " + email_from

                    # Only get email body from Eakins emails
                    if (email_subject == EAKINS_SUBJECT and email_from == EAKINS_SENDER):

                        # Extract body from email
                        payload = msg.get_payload()
                        body = extract_body(payload)

                        # Extract shift start and end times as strings
                        shift_start = re.search("Shift start: (.*) \| Shift end: ", body)
                        shift_end = re.search("Shift end: (.*) Additional", body)

                        # Check to see if date fits preferences
                        if (date_check(shift_start)):

                            # Print if date checking is successfull
                            print '\n' + time.strftime("%Y-%m-%d %H:%M:%S") + '\t' +"SHIFT ACQUIRED" + '\n'
                            print "Shift start:     " + shift_start.group(1)
                            print "Shift end:       " + shift_end.group(1) + '\n'

                            # Get URL for shift and take it
                            take_shift = re.search("<a href ='(.*)' style = 'font-size: 20px; color: green;'>I'll Take It</a>", body)
                            webbrowser.open(take_shift.group(1), new=0, autoraise=True)

                            # New shift taken
                            return True

        # Print update every time we loop around the last X emails
        print time.strftime("%Y-%m-%d %H:%M:%S") + '\t' + "Searched last 3 emails, no new shifts found."

        # No new shifts taken
        return False

    except Exception, e:
        print str(e)


# Continually return the most recent email
shift_taken = False
while (not shift_taken):
    shift_taken = read_email()
