import pyzmail
import imaplib
import bs4
import getpass
import webbrowser
import re
import sys

'''List of accepted service providers and respective imap link'''
servers = [('Gmail', 'imap.gmail.com'), ('Outlook', 'imap-mail.outlook.com'),
           ('Hotmail', 'imap-mail.outlook.com'), ('Yahoo', 'imap.mail.yahoo.com'),
           ('ATT', 'imap.mail.att.net'), ('Comcast', 'imap.comcast.net'),
           ('Verizon', 'incoming.verizon.net'), ('AOL', 'imap.aol.com'),
           ('Zoho', 'imap.zoho.com')]

# add to words if more words found
'''Key words for unsubscribe link - add more if found'''
words = ['unsubscribe', 'subscription', 'optout']


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
        print('This program searches your email for junk mail to unsubscribe from and delete')
        print('Supported emails: Gmail, Outlook, Hotmail, Yahoo, AOL, Zoho,')
        print('AT&T, Comcast, and Verizon')
        print('Please note: you may need to allow access to less secure apps')
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
                print('\nNo usable email type detected, try a different account')
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
            print(f'\nAn error occurred while attempting to log in: {e}\n')
            return False

    '''Attempt to log in to server. On failure, force user to re-enter info'''

    def access_server(self, readonly=True):
        if self.email == '':
            self.get_info()
        attempt = self.login(readonly)
        if not attempt:
            self.access_server(readonly)

    '''Search for emails with unsubscribe in the body. If sender not already in
    senderList, parse email for unsubscribe link. If link found, add name, email,
    link (plus metadata for decisions) to senderList. If not, add to noLinkList.
    '''

    def get_emails(self):
        print('Getting emails with unsubscribe in the body\n')
        _, ui_ds = self.imap.search(None, 'BODY', 'unsubscribe')

        print('Getting links and addresses\n')
        for UID in ui_ds[0].split():
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
                        print("Error decoding HTML payload. Skipping this email.")
                        continue

                    soup = bs4.BeautifulSoup(html, 'html.parser')
                    elems = soup.select('a')

                    '''For each anchor tag, use regex to search for key words'''
                    for elem in elems:
                        for word in self.wordCheck:
                            if word.search(str(elem)):
                                print('Link found')
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
                    print('No link found')
                    not_in_list = True
                    for noLinkers in self.noLinkList:
                        if sender[0][1] in noLinkers:
                            not_in_list = False
                    if not_in_list:
                        self.noLinkList.append([sender[0][0], sender[0][1]])

        try:
            print('\nLogging out of email server\n')
            self.imap.logout()
        except Exception as e:
            print(f'Error logging out: {e}')

    def display_email_info(self):
        if self.noLinkList:
            print('Could not find unsubscribe links from these senders:')
            no_list = '| '
            for i in range(len(self.noLinkList)):
                no_list += (str(self.noLinkList[i][0])+' | ')
            print(no_list)
        if self.senderList:
            print('\nFound unsubscribe links from these senders:')
            full_list = '| '
            for i in range(len(self.senderList)):
                full_list += (str(self.senderList[i][0])+' | ')
            print(full_list)

    def decisions(self):
        def choice(user_input):
            if user_input.lower() == 'y':
                return True
            elif user_input.lower() == 'n':
                return False
            else:
                return None

        self.display_email_info()
        print('\nYou may now decide which emails to unsubscribe from and/or delete')
        print('Navigating to unsubscribe links may not automatically unsubscribe you')
        print('Please note: deleted emails cannot be recovered\n')
        for j in range(len(self.senderList)):
            while True:
                # unsub = input('Open unsubscribe link from ' + str(self.senderList[j][0]) + ' (Y/N): ')
                # c = choice(unsub)
                c = True  # Open all unsubscribe links
                if c:
                    self.senderList[j][3] = True
                    self.goToLinks = True
                    break
                elif not c:
                    break
                else:
                    print('Invalid choice, please enter \'Y\' or \'N\'.\n')
            while True:
                # delete = input('Delete emails from ' + str(self.senderList[j][1]) + ' (Y/N): ')
                # d = choice(delete)
                d = False  # Do not delete e-mails
                if d:
                    self.senderList[j][4] = True
                    self.delEmails = True
                    break
                elif not d:
                    break
                else:
                    print('Invalid choice, please enter \'Y\' or \'N\'.\n')

    '''Navigate to selected unsubscribe, 10 at a time'''

    def open_links(self):
        if not self.goToLinks:
            print('\nNo unsubscribe links selected to navigate to')
        else:
            print('\nUnsubscribe links will be opened 10 at a time')
            counter = 0
            filtered_sender_list = list(dict.fromkeys(self.senderList))  # remove duplicates from list
            for i in range(len(filtered_sender_list)):
                if filtered_sender_list[i][3]:
                    webbrowser.open(filtered_sender_list[i][2])
                    counter += 1
                    if counter == 10:
                        print('Navigating to unsubscribe links')
                        input('Press \'Enter\' to continue: ')
                        counter = 0

    '''Log back into IMAP servers, NOT in readonly mode, and delete emails from
    selected providers. Note: only deleting emails with unsubscribe in the body.
    Emails from provider without unsubscribe in the body will not be deleted.
    '''

    def delete_emails(self):
        if not self.delEmails:
            print('\nNo emails selected to delete')
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
            print('\nLogging out of email server')
            self.imap.logout()

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
        self.get_emails()
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
