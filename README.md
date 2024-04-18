# AutoUnsubscriber

> This program is an email auto-unsubscriber. Depending on your email provider and settings, it may require you to allow access to less secure apps.
> 
> It uses IMAP to log into your email.
From there, it goes through every email with "unsubscribe" in the body, parses the HTML, and uses regex to search through anchor tags for keywords that indicate an unsubscribe link (unsubscribe, optout, etc.).
>
> If it finds a match, it grabs the href link and puts the address and link in a list.
After the program has a list of emails and links, for each address in the list, it gives the user the option to navigate to the unsubscribe link and to delete emails with unsubscribe in the body from the sender.
Once the program finishes going through the list, it gives the user the option to run the program again on the same email address, run it on a different email address, or quit the program.

### Prerequisites 
> If you have 2-Step verification enabled in Gmail, please follow the steps in this [link](https://support.google.com/accounts/answer/185833)
> 
> The password generated will be used for signing in instead of your original password.

### Run
```bash
git clone https://github.com/kbkozlev/AutoUnsubscriber.git
cd AutoUnsubscriber
pip install -r requirements.txt
python AutoUnsubscriber.py
```