
import copy
import logging

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



class arc_o365(object):

    def __init__(self, config, token_filename, scopes=_scopes_default, add_scopes=None, **kwargs):
        """ Initialized the ms graph API.

            config -- a dict of the configuration parameters
            config_static -- the statically configured elements; only used to detect misconfiguration of security critical values that should not be checked into source control
            scopes -- the security scopes that will be requested
            token_filename -- the file used to store the security token

        """

        credentials = (config.CLIENT_ID, config.CLIENT_SECRET)

        if add_scopes is not None:
            scopes = copy.copy(scopes)
            scopes.extend(add_scopes)

        token_backend = O365.FileSystemTokenBackend(token_path='.', token_filename=token_filename)
        #log.info("before account object creation")
        account = O365.Account(credentials, scopes=scopes, token_backend=token_backend, **kwargs)
        #log.info(f"after account object creation: { account }")


        if not account.is_authenticated:
            log.info(f"Authenticating account associated with file { token_filename }")
            account.authenticate()
            if not account.is_authenticated:
                log.fatal(f"Cannot authenticate account")
                raise Exception("Could not authenticate with MS Graph API")

        self.account = account

    def get_account(self):
        return self.account

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
