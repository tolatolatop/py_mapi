import datetime
import sys

import win32com
import win32com.client

f = sys.stdout


def get_outlook():
    res = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    return res


def get_accounts():
    res = win32com.client.Dispatch("Outlook.Application").Session.Accounts
    return res


class MailFolder(object):

    def __init__(self, path: str, root):
        self.root = root
        self.path = path.replace(' (仅限于此计算机)', '')
        self.parent = None
        self.obj = None if not self.is_root() else root

    def walk(self):
        folders, mails = self.list()
        folders = list(folders)
        yield iter(folders), mails
        for folder in folders:
            for sub_folders, sub_mails in folder.walk():
                yield sub_folders, sub_mails

    def is_root(self):
        return self.path == '/'

    def __exists(self):
        if self.is_root() or self.obj:
            return True

        if self.parent is None:
            parent_path = '/' + '/'.join(self.path.split('/')[1:-1])
            self.parent = MailFolder(parent_path, self.root)

        if isinstance(self.parent, MailFolder):
            folders, _ = self.parent.list()
            for folder in folders:
                if folder == self:
                    self.obj = folder.obj
                    return True
            else:
                raise FileNotFoundError(self)

    def list_folder(self):
        if not self.__exists():
            raise
        for obj in self.obj.folders:
            if self.is_root():
                path = '/' + str(obj)
            else:
                path = self.path + '/' + str(obj)
            folder = MailFolder(path, self.root)
            folder.parent = self
            folder.obj = obj
            yield folder

    def list_mail(self):
        if not self.__exists():
            raise
        for item in self.obj.Items:
            mail = Mail(self, item)
            yield mail

    def list(self):
        if self.__exists():
            folders = self.list_folder()
            mails = self.list_mail()
            return folders, mails
        else:
            raise

    def __str__(self):
        return self.path

    def __eq__(self, other):
        return self.root == other.root and self.path == other.path


class Mail(object):

    def __init__(self, folder, obj):
        self.folder = folder
        self.obj = obj

    @property
    def html(self):
        return self.obj.HTMLBody

    @property
    def sender_address(self):
        return self.obj.SenderEmailAddress

    @property
    def subject(self):
        return self.obj.Subject

    @property
    def received_time(self) -> datetime.datetime:
        return self.obj.ReceivedTime


if __name__ == '__main__':
    outlook = get_outlook()
    accounts = get_accounts()
    for account in accounts:
        inbox = outlook.Folders(account.DeliveryStore.DisplayName)
        root_mail_box = MailFolder('/收件箱', inbox)
        for mail_folder, mails in root_mail_box.walk():
            for mail in mails:
                print(mail.received_time)
