#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Microsoft AutoUpdate Command Line Utility."""

import os
import sys
import httplib
import plistlib
import tempfile
import subprocess


UUID = 'C1297A47-86C4-4C1F-97FA-950631F94777'
NAMES = {
    'XCEL15': 'Excel',
    'ONMC15': 'OneNote',
    'OPIM15': 'Outlook',
    'PPT315': 'PowerPoint',
    'MSWD15': 'Word',
    'MSau03': 'AutoUpdate',
    'SLVT': 'Silverlight',
}


BASEURL = '/pr/%s/OfficeMac/' % UUID
PREF_PATH = 'Library/Preferences/com.microsoft.autoupdate2.plist'


def get_plist(path):
    """Return plist dict regardless of format."""
    plist = subprocess.check_output(['/usr/bin/plutil',
                                     '-convert', 'xml1',
                                     path, '-o', '-'])
    return plistlib.readPlistFromString(plist)


def set_pref(k='HowToCheck', v='Manual'):
    """Set AU pref k to v."""
    subprocess.call(['/usr/bin/defaults',
                     'write',
                     'com.microsoft.autoupdate2',
                     k,
                     v])


def check(home=None, pref=PREF_PATH):
    """Collect info about available updates."""
    results = []

    if home is None:
        home = os.getenv('HOME')

    pref = os.path.join(home, pref)

    try:
        plist = get_plist(os.path.expanduser(pref))
    except Exception as e:
        raise Exception('Failed to read MAU prefs: %s' % e)

    apps = plist.get('Applications')

    for a in apps.items():
        path, info = a
        app_id = info.get('Application ID')

        # Skip Microsoft Error Reporting
        if app_id == 'Merp2':
            continue

        try:
            app_info = get_plist(os.path.join(path, 'Contents/Info.plist'))
        except Exception as e:
            continue

        result = {'id': app_id, 'installed': app_info.get('CFBundleVersion')}
        result['name'] = NAMES.get(app_id, app_id)
        result['lcid'] = info.get('LCID')

        # Lync (UCCP14) is special
        filename = '0409%s.xml' if app_id == 'UCCP14' else '0409%s-chk.xml'

        try:
            conn = httplib.HTTPSConnection('officecdn.microsoft.com')
            conn.request("GET", BASEURL + filename % app_id)
            data = conn.getresponse().read()
        except Exception as e:
            raise Exception('Failed to check for updates: %s' % e)

        try:
            p = plistlib.readPlistFromString(data)

            if app_id == 'UCCP14':  # Lync being special again
                p = p[0]
                result['location'] = p.get('Location')
                versions = p['Triggers']['Lync']['Versions']
                result['needs_update'] = result['installed'] in versions
            else:
                result['date'] = p.get('Date')
                result['type'] = p.get('Type')
                result['available'] = p.get('Update Version')
                result['needs_update'] = result['installed'] != result['available']

                if result['needs_update']:
                    url = BASEURL + filename % app_id
                    conn.request("GET", url.replace('-chk', ''))
                    data = conn.getresponse().read()
                    # Fetch the update details
                    updates = plistlib.readPlistFromString(data)
                    for i in updates:
                        if i.get('Baseline Version') == result['installed'] or app_id == 'MSau03':
                            result['location'] = i.get('Location')
                            result['size'] = i.get('FullUpdaterSize')
                            results.append(result)
        except Exception as e:
            print('Failed to check %s' % app_id, e)

        finally:
            conn.close()

    return results


def download(url):
    """Download url to temporary folder."""
    temp = os.path.join(tempfile.gettempdir(), 'mowgli')

    if not os.path.exists(temp):
        os.mkdir(temp)

    fn = url.split('/')[-1]
    fp = os.path.join(temp, fn)

    if not os.path.exists(fp):
        subprocess.call(['/usr/bin/curl', url, '-o', fp])

    return fp


def install(pkg):
    """Install a package."""
    r = subprocess.call(['/usr/sbin/installer', '-pkg', pkg, '-target', '/'])
    if r > 0:
        raise Exception('Failed to install package %s' % pkg)

    os.remove(pkg)


if __name__ == '__main__':

    if len(sys.argv) < 2:
        print("usage: {0} [-l | -i [home]".format(os.path.basename(sys.argv[0])))
        sys.exit(1)

    if sys.argv[1] == 'enable':
        set_pref('HowToCheck', 'Manual')
        print('* Automatic updates enabled')
        sys.exit(0)

    if sys.argv[1] == 'disable':
        set_pref('HowToCheck', 'Automatic')
        print('* Automatic updates disabled')
        sys.exit(0)

    print("* Finding available software")

    if len(sys.argv) > 2:
        if not os.path.exists(sys.argv[2]):
            raise Exception('Invalid MAU pref file %s' % sys.argv[2])
        updates = check(sys.argv[2])
    else:
        updates = check()

    updates = [u for u in updates if u['needs_update']]

    if sys.argv[1] == '-l':
        for u in updates:
            print("  * {id}\n       {name} {installed} - {available} [{type}]".format(**u))

    if sys.argv[1] == '-ia':
        for u in updates:
            print('* Downloading %s %s' % (u['name'], u['available']))
            pkg = download(u['location'])
            print('* Installing %s' % pkg)
            install(pkg)

    if len(updates) < 1:
        print("* No updates available")
        sys.exit(1)

    sys.exit(0)
