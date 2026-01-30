
import copy
import logging
import datetime
import re
import base64

#import O365_local as O365
import O365

#import init_logging
log = logging.getLogger(__name__)

_scopes_default = [
        'https://graph.microsoft.com/Files.ReadWrite.All',
        'https://graph.microsoft.com/Mail.Read',
        'https://graph.microsoft.com/Mail.Read.Shared',
        'https://graph.microsoft.com/Mail.Send',
        'https://graph.microsoft.com/Mail.Send.Shared',
        'https://graph.microsoft.com/offline_access',
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/User.ReadBasic.All',
        'https://graph.microsoft.com/Contacts.ReadWrite',
        'https://graph.microsoft.com/Contacts.ReadWrite.Shared',
        #'https://graph.microsoft.com/GroupMember.ReadWrite.All',   # needs admin auth
        #'https://graph.microsoft.com/Group.ReadWrite.All',         # needs admin auth
        #'https://graph.microsoft.com/Directory.AccessAsUser.All',  # needs admin auth

        #'https://microsoft.sharepoint-df.com/AllSites.Read',
        #'https://microsoft.sharepoint-df.com/MyFiles.Read',
        #'https://microsoft.sharepoint-df.com/MyFiles.Write',
        'https://graph.microsoft.com/Sites.ReadWrite.All',
        'basic',
        ]

_token_filename_default = "o365_token.txt"


MATCH_TO_FIRST_UNDERSCORE = re.compile(r"^([^_]+)_")

class arc_o365(object):

    def __init__(self, config, token_filename, scopes=None, add_scopes=None, **kwargs):
        """ Initialized the ms graph API.

            config -- a dict of the configuration parameters
            token_filename -- the file used to store the security token
            scopes -- the security scopes that will be requested
            add_scopes -- additional scopes to add to request

        """
        credentials = (config.CLIENT_ID, config.CLIENT_SECRET)

        if scopes is None:
            log.info("arc_o365: scopes was None; using defaults")
            scopes = _scopes_default
        else:
            log.info(f"arc_o365: scopes was { scopes }")


        if add_scopes is not None:
            scopes = copy.copy(scopes)
            scopes.extend(add_scopes)

        token_backend = O365.FileSystemTokenBackend(token_path='.', token_filename=token_filename)
        #log.info("before account object creation")
        account = O365.Account(credentials, token_backend=token_backend, **kwargs)
        #log.info(f"after account object creation: { account }")


        if not account.is_authenticated:
            log.info(f"Authenticating account associated with file { token_filename }.  Scopes { scopes }")
            account.authenticate(requested_scopes=scopes)
            if not account.is_authenticated:
                log.fatal(f"Cannot authenticate account")
                raise Exception("Could not authenticate with MS Graph API")

        self.account = account
        self.config = config

    def get_account(self):
        return self.account


    # search the specified email address for mail matching the specified subject pattern.
    def search_mail(self, email_address, subj_pattern, limit=1):

        mailbox = self.get_account().mailbox(resource=email_address)
        builder = mailbox.new_query()
        dt = datetime.datetime(1900, 1, 1)
        matcher = builder.chain_and(
                builder.greater('sentDateTime', dt),
                builder.contains("subject", subj_pattern))

        messages = mailbox.get_messages(query=matcher, order_by="sentDateTime desc", limit=limit, download_attachments=True)
        message_list = list(messages)

        if len(message_list) == 0:
            log.debug(f"No messages found")
        else:
            log.debug(f"found { len(message_list) } messages")

        return message_list


    # fetch workforce reports from the specified shared mailbox
    def fetch_workforce_reports(self, dro_id, limit=1):

        message_match_string = f"DR { dro_id } Automated Workforce Reports"
        message_list = self.search_mail(self.config.PROGRAM_EMAIL, message_match_string, limit=1)

        if len(message_list) == 0:
            error = f"Could not find an email that matches '{ message_match_string }'"
            log.fatal(error)
            raise(Exception(error))


        return_list = []
        for message in message_list:
            attach_dict = {}
            # read the attachments
            for attachment in message.attachments:
                attach_name = attachment.name
                name_type = attach_name
                name_before_underscore = MATCH_TO_FIRST_UNDERSCORE.match(attach_name)

                if name_before_underscore is not None:
                    name_type  = name_before_underscore.group(1)

                #log.debug(f"attachment { attachment.name } name_type { name_type } size { attachment.size }")
                attach_dict[name_type] = base64.b64decode(attachment.content)

            attach_dict['subject'] = message.subject
            return_list.append(attach_dict)

        if limit == 1:
            if len(message_list) > 0:
                return return_list[0]
            else:
                return None
        else:
            return return_list




import dotenv
def main():
    global log
    #log = logging.getLogger(__name__)
    class AttrDict(dict):
        def __init__(self, *args, **kwargs):
            super(AttrDict, self).__init__(*args, **kwargs)
            self.__dict__ = self

    config = AttrDict()
    config_dotenv = dotenv.dotenv_values(dotenv_path='.env_test', verbose=True)
    for key, value in config_dotenv.items():
        config[key] = value

    config_static = AttrDict()
    #log.info(f"before arc_o365 call")
    o365 = arc_o365(config)
    #log.info(f"after arc_o365 call")

if __name__ == "__main__":
    init_logging.init()
    main()
