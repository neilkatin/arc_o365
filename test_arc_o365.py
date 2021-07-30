

import logging

import pytest
import arc_o365
import dotenv
import init_logging

log = logging.getLogger(__name__)

def init_account(add_scopes=None):
    class AttrDict(dict):
        def __init__(self, *args, **kwargs):
            super(AttrDict, self).__init__(*args, **kwargs)
            self.__dict__ = self

    config = AttrDict()
    config_dotenv = dotenv.dotenv_values(dotenv_path='.env_test', verbose=True)
    for key, value in config_dotenv.items():
        config[key] = value

    config_static = AttrDict()

    o365 = arc_o365.arc_o365(config, config_static, add_scopes=add_scopes, token_filename="test_token.txt")
    account = o365.get_account()
    assert account is not None

    return account

def xtest_sharepoint():

    account = init_account()

    sharepoint = account.sharepoint()
    assert sharepoint is not None

    root = sharepoint.get_root_site()
    assert root is not None

    subsites = root.get_subsites()
    log.info(f"root: there are { len(subsites) } subsites")

    for site in subsites:
        log.info(f"root subsite: { site }")

    arc_site = sharepoint.get_site('americanredcross.sharepoint.com')
    subsites = arc_site.get_subsites()
    log.info(f"arc: there are { len(subsites) } subsites")
    for site in subsites:
        log.info(f"arc subsite: { site }")

    results = sharepoint.search_site("NCCR")
    log.info(f"search: there are { len(subsites) } subsites")
    for site in results:
        log.info(f"search subsite: { site }")


def xxxtest_groups():
    account = init_account()

    extra_scopes= [ 'GroupMember.ReadWrite.All', 'Group.ReadWrite.All' ]

    groups = account.groups()
    log.info(f"groups: got { groups }")

    teams = groups.get_joined_teams()


def test_contacts():
    log.info(f"test_contacts: called")
    account = init_account()

    ab_dir = account.address_book()
    assert ab_dir is not None

    folders = ab_dir.get_folders()
    assert folders is not None
    assert isinstance(folders, list)

    log.info(f"len of folders is { len(folders) }")

    count = 0
    for f in folders:
        log.info(f"got returned ab folder entry { f }")
        count += 1
        if count > 100:
            break


