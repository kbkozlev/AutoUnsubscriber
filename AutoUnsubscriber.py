import pyzmail
import imaplib
import bs4
import getpass
import webbrowser
import re
import sys
from advancedprinter import print

'''List of accepted service providers and respective imap link'''
servers = [('Gmail', 'imap.gmail.com'), ('Outlook', 'imap-mail.outlook.com'),
           ('Hotmail', 'imap-mail.outlook.com'), ('Yahoo', 'imap.mail.yahoo.com'),
           ('ATT', 'imap.mail.att.net'), ('Comcast', 'imap.comcast.net'),
           ('Verizon', 'incoming.verizon.net'), ('AOL', 'imap.aol.com'),
           ('Zoho', 'imap.zoho.com')]

# add to words if more words found
'''Key words for unsubscribe link - add more if found'''
words = ['unsubscribe', 'subscription', 'optout', 'here']  # added 'here' as some links are using this keyword


class AutoUnsubscriber:

    def __init__(self):
        self.email = ''
        self.user = None
        self.password = ''
        self.imap = None
        self.goToLinks = False
        self.delEmails = False
        self.senderList = []
        self.noLinkList = []
        self.wordCheck = []
        self.providers = []
        for i in range(len(servers)):
            self.providers.append(re.compile(servers[i][0], re.I))
        for i in range(len(words)):
            self.wordCheck.append(re.compile(words[i], re.I))

    '''Get initial user info - email, password, and service provider'''

    def get_info(self):
        print('This program searches your email for junk mail to unsubscribe from and delete\n'
              'Supported emails: Gmail, Outlook, Hotmail, Yahoo, AOL, Zoho, AT&T, Comcast, and Verizon', c='cyan')
        print('Please note: you may need to allow access to less secure apps', c='yellow')
        get_email = True
        while get_email:
            self.email = input('\nEnter your email address: ')
            for j in range(len(self.providers)):
                choice = self.providers[j].search(self.email)
                if choice is not None:
                    self.user = servers[j]
                    print('\nLooks like you\'re using a ' + self.user[0] + ' account\n')
                    get_email = False
                    break
            if self.user is None:
                print('\nNo usable email type detected, try a different account', c='red')
        self.password = getpass.getpass('Enter password for ' + self.email + ': ')

    '''Log in to IMAP server, argument determines whether readonly or not'''

    def login(self, read=True):
        try:
            self.imap = imaplib.IMAP4_SSL(self.user[1])
            self.imap.login(self.email, self.password)
            self.imap.select('inbox', readonly=read)
            print('\nLog in successful\n')
            return True
        except Exception as e:
            print(f'\nAn error occurred while attempting to log in:', c='red')
            print(f"{e}\n")
            return False

    '''Attempt to log in to server. On failure, force user to re-enter info'''

    def access_server(self, readonly=True):
        if self.email == '':
            self.get_info()
        attempt_count = 0
        while attempt_count < 3:
            attempt = self.login(readonly)
            if attempt:
                return
            attempt_count += 1
        print("Maximum attempts reached. Exiting.", c='yellow')
        exit()

    '''Search for emails with unsubscribe in the body. If sender not already in
    senderList, parse email for unsubscribe link. If link found, add name, email,
    link (plus metadata for decisions) to senderList. If not, add to noLinkList.
    '''

    def get_emails(self, scan_limit=None):  # Scan limit added for testing purposes
        print('Getting emails with unsubscribe in the body\n')
        _, ui_ds = self.imap.search(None, 'BODY', 'unsubscribe')

        print('Getting links and addresses\n')
        email_count = 0  # Initialize email counter
        for UID in ui_ds[0].split():
            if scan_limit is not None and email_count >= scan_limit:
                break  # Stop processing emails if scan limit reached
            _, data = self.imap.fetch(UID, '(RFC822)')
            raw = data[0][1]

            '''Get address and check if sender already in senderList'''
            msg = pyzmail.PyzMessage.factory(raw)
            sender = msg.get_addresses('from')
            try_sender = True
            for spammers in self.senderList:
                if sender[0][1] in spammers:
                    try_sender = False

            '''If not, search for link'''
            if try_sender:
                '''Encode and decode to cp437 to handle unicode errors and get
                rid of characters that can't be printed by Windows command line
                which has default setting of cp437
                '''
                sender_name = sender[0][0].encode('cp437', 'ignore').decode('cp437')
                print('Searching for unsubscribe link from ' + sender_name)
                url = False

                '''Parse html for elements with anchor tags'''
                if msg.html_part is not None:
                    try:
                        html = msg.html_part.get_payload().decode('utf-8')
                    except UnicodeDecodeError:
                        print("Error decoding HTML payload. Skipping this email.", c='red')
                        continue

                    soup = bs4.BeautifulSoup(html, 'html.parser')
                    elems = soup.select('a')

                    '''For each anchor tag, use regex to search for key words'''
                    for elem in elems:
                        for word in self.wordCheck:
                            if word.search(str(elem)):
                                print('Link found', c='green')
                                url = elem.get('href')
                                break
                        if url:
                            break

                '''If link found, add info to senderList
                format: (Name, email, link, go to link, delete emails)
                If no link found, add to noLinkList
                '''
                if url:
                    self.senderList.append([sender_name, sender[0][1], url, False, False])
                else:
                    print('No link found', c='red')
                    not_in_list = True
                    for noLinkers in self.noLinkList:
                        if sender[0][1] in noLinkers:
                            not_in_list = False
                    if not_in_list:
                        self.noLinkList.append([sender[0][0], sender[0][1]])

            email_count += 1  # Increment email counter after processing each email

        try:
            print('\nLogging out of email server\n', c='yellow')
            self.imap.logout()
        except Exception as e:
            print(f'Error logging out: {e}', c='red')

    def display_email_info(self):
        if self.noLinkList:
            print('Could not find unsubscribe links from these senders:', c='red')
            no_list = '| '
            for i in range(len(self.noLinkList)):
                no_list += (str(self.noLinkList[i][0])+' | ')
            print(no_list)
        if self.senderList:
            print('\nFound unsubscribe links from these senders:', c='green')
            full_list = '| '
            for i in range(len(self.senderList)):
                full_list += (str(self.senderList[i][0])+' | ')
            print(full_list)

    def decisions(self):
        def get_choice(prompt):
            while True:
                user_input = input(prompt).lower()
                if user_input == 'y':
                    return True
                elif user_input == 'n':
                    return False
                else:
                    print('Invalid choice, please enter \'Y\' or \'N\'.', c='yellow')

        self.display_email_info()
        print('\nYou may now decide which emails to unsubscribe from and/or delete', c='cyan')
        print('Navigating to unsubscribe links may not automatically unsubscribe you', c='cyan')
        print('Please note: deleted emails cannot be recovered\n', c='yellow')

        choice = input("Do you want to decide for all e-mails or separate (A/S): ").lower()
        if choice == 'a':
            unsub = get_choice("Unsubscribe from all (Y/N): ")
            delete = get_choice("Delete all (Y/N): ")

            if unsub:
                self.senderList = [[sender[0], sender[1], sender[2], True, sender[4]] for sender in self.senderList]
                self.goToLinks = True
            if delete:
                self.senderList = [[sender[0], sender[1], sender[2], sender[3], True] for sender in self.senderList]
                self.delEmails = True
        else:
            for sender in self.senderList:
                unsub = get_choice('Open unsubscribe link from {} (Y/N): '.format(sender[0]))
                if unsub:
                    sender[3] = True
                    self.goToLinks = True

                delete = get_choice('Delete emails from {} (Y/N): '.format(sender[1]))
                if delete:
                    sender[4] = True
                    self.delEmails = True

    '''Navigate to selected unsubscribe, 20 at a time'''

    def open_links(self):
        if not self.goToLinks:
            print('\nNo unsubscribe links selected to navigate to', c='yellow')
        else:
            seen_emails = set()  # Initialize a set to store unique email addresses
            filtered_sender_list = []  # Initialize an empty list to store filtered sender information

            # Iterate over self.senderList and add unique elements to filtered_sender_list
            for sender_info in self.senderList:
                email = sender_info[1]
                if email not in seen_emails:
                    filtered_sender_list.append(sender_info)
                    seen_emails.add(email)

            total_entries = len(filtered_sender_list)  # Calculate total number of entries
            print(f"\nTotal number of entries: {total_entries}")
            print('\nUnsubscribe links will be opened 20 at a time')
            counter = 0
            for sender_info in filtered_sender_list:
                if sender_info[3]:
                    webbrowser.open(sender_info[2])
                    counter += 1
                    if counter == 20:  # Increased to 20 to speed up process
                        print(f'{counter} unsubscribe links navigated to. {total_entries - counter} entries left.')
                        input('Press \'Enter\' to continue: ')
                        counter = 0

        try:
            print('\nLogging out of email server\n', c='yellow')
            self.imap.logout()
        except Exception as e:
            print(f'Error logging out: {e}', c='red')

    '''Log back into IMAP servers, NOT in readonly mode, and delete emails from
    selected providers. Note: only deleting emails with unsubscribe in the body.
    Emails from provider without unsubscribe in the body will not be deleted.
    '''

    def delete_emails(self):
        if not self.delEmails:
            print('\nNo emails selected to delete', c='yellow')
        else:
            print('\nLogging into email server to delete emails')
            '''Pass false to self.login() so as to NOT be in readonly mode'''
            self.login(False)
            del_total = 0
            for i in range(len(self.senderList)):
                if self.senderList[i][4]:
                    print('Searching for emails to delete from ' + str(self.senderList[i][1]))
                    from_sender = 'FROM ' + str(self.senderList[i][1])
                    '''Search for unsubscribe in body from selected providers'''
                    _, del_ui_ds = self.imap.search(None, 'BODY', 'unsubscribe', from_sender)
                    del_count = 0
                    for DelUID in del_ui_ds[0].split():
                        '''Delete emails from selected providers'''
                        self.imap.store(DelUID, '+FLAGS', '\\Deleted')
                        self.imap.expunge()
                        del_count += 1
                    print('Deleted ' + str(del_count) + ' emails from ' + str(self.senderList[i][1]))
                    del_total += del_count
            print('\nTotal emails deleted: ' + str(del_total))
            try:
                print('\nLogging out of email server\n', c='yellow')
                self.imap.logout()
            except Exception as e:
                print(f'Error logging out: {e}', c='red')

    '''For re-running on same email. Clear lists, reset flags, but use same info
    for email, password, email provider, etc.
    '''

    def run_again(self):
        self.goToLinks = False
        self.delEmails = False
        self.senderList = []
        self.noLinkList = []

    '''Reset everything to get completely new user info'''

    def new_email(self):
        self.email = ''
        self.user = None
        self.password = ''
        self.imap = None
        self.run_again()

    '''Called after program has run, allow user to run again on same email, run
    on a different email, or quit the program
    '''

    def next_move(self):
        print('\nRun this program again on the same email, a different email, or quit?\n')
        while True:
            print('Press \'A\' to run again on ' + str(self.email))
            print('Press \'D\' to run on a different email address')
            again = input('Press \'Q\' to quit: ')
            if again.lower() == 'a':
                print('\nRunning program again for ' + str(self.email) + '\n')
                self.run_again()
                return True
            elif again.lower() == 'd':
                print('\nPreparing program to run on a different email address\n')
                self.new_email()
                return False
            elif again.lower() == 'q':
                print('\nSo long, space cowboy!\n')
                sys.exit()
            else:
                print('\nInvalid choice, please enter \'A\', \'D\' or \'Q\'.\n')

    '''Full set of program commands. Works whether it has user info or not'''

    def full_process(self):
        self.access_server()
        self.get_emails(scan_limit=100)  # set limit for testing
        if self.senderList:
            self.decisions()
            self.open_links()
            self.delete_emails()
        else:
            print('No unsubscribe links detected')

    '''Loop to run program and not quit until told to by user or closed'''

    def usage_loop(self):
        self.full_process()
        while True:
            self.next_move()
            self.full_process()


def main():
    auto = AutoUnsubscriber()
    auto.usage_loop()


if __name__ == '__main__':
    main()
