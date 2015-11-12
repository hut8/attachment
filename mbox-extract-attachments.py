#!/usr/bin/env python
# -*- coding: utf-8 -*-
# mbox-extract-attachments.py - Extract attachments from mbox files
# 16/March/2012
# Copyright (C) 2012 Pablo Castellano <pablo@anche.no>
#
# Adapted by Liam Bowen <LiamBowen@gmail.com>
# November 2015
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.
#

# Notes (RFC 1341):
# The use of a Content-Type of multipart in a body part within another multipart entity is explicitly allowed. In such cases, for obvious reasons, care must be taken to ensure that each nested multipart entity must use a different boundary delimiter. See Appendix C for an example of nested multipart entities.
# The use of the multipart Content-Type with only a single body part may be useful in certain contexts, and is explicitly permitted.
# The only mandatory parameter for the multipart Content-Type is the boundary parameter, which consists of 1 to 70 characters from a set of characters known to be very robust through email gateways, and NOT ending with white space. (If a boundary appears to end with white space, the white space must be presumed to have been added by a gateway, and should be deleted.) It is formally specified by the following BNF

# Related RFCs: 2047, 2044, 1522


__author__ = "Pablo Castellano <pablo@anche.no>"
__author__ = ', '.join([__author__, "Liam Bowen <liambowen@gmail.com>"])
__license__ = "GNU GPLv3+"
__version__ = 1.3
__date__ = "2015-11-12"

import mailbox
import base64
import os
import sys
import email
try:
    from tqdm import tqdm
except ImportError:
    tqdm = list
    print "progress bar library not found"
    print "run: pip install tqdm"

BLACKLIST = ('signature.asc', 'message-footer.txt', 'smime.p7s')
VERBOSE = 1

attachments = 0  # Count extracted attachment
skipped = 0

# Search for filename or find recursively if it's multipart
def extract_attachment(payload):
    global attachments, skipped
    filename = payload.get_filename()

    # Recursive Case
    if filename is None:
        if payload.is_multipart():
            for payl in payload.get_payload():
                extract_attachment(payl)
        return

    # Base Case
    if filename.find('=?') != -1:
        ll = email.header.decode_header(filename)
        filename = ""
        for l in ll:
            filename = filename + l[0]

    if filename in BLACKLIST:
        skipped += 1
        return

    # Can the filename not be specified?
    # if filename is None:
    #     filename = "unknown_%d_%d.txt" % (i, p)

    content = payload.as_string()
    # Skip headers, go to the content
    fh = content.find('\n\n')
    content = content[fh:]

    # if it's base64....
    if payload.get('Content-Transfer-Encoding') == 'base64':
        content = base64.decodestring(content)
        # TODO: Handle this:
        # File "/usr/lib64/python2.6/base64.py", line 321, in decodestring
        #     return binascii.a2b_base64(s)
        # binascii.Error: Incorrect padding

    # quoted-printable
    # what else? ...

    n = 1
    filename = orig_filename = os.path.basename(filename)
    while os.path.exists(filename):
        # TODO: Detect if it's just a duplicate
        filename = orig_filename + "." + str(n)
        n += 1

    fp = None
    try:
        fp = open(filename, "w")
        fp.write(content)
    except IOError as e:
        print "io error while saving attachment: %s" % str(e)
        return
    finally:
        if fp is not None: fp.close()

    attachments = attachments + 1

def main(filename, directory):
    mb = mailbox.mbox(filename)
    print "counting messages... "
    message_count = len(mb)
    print "done (found %s)" % message_count

    os.chdir(directory)

    # tqdm = list
    for i in tqdm(range(message_count), ascii=True):
        mes = mb.get_message(i)
        em = email.message_from_string(mes.as_string())

        # subject = em.get('Subject')
        # if subject is not None and subject.find('=?') != -1:
        #     ll = email.header.decode_header(subject)
        #     subject = ""
        #     for l in ll:
        #         subject = subject + l[0]

        # em_from = em.get('From')
        # if em_from is not None and em_from.find('=?') != -1:
        #     ll = email.header.decode_header(em_from)
        #     em_from = ""
        #     for l in ll:
        #         em_from = em_from + l[0]

        if em.is_multipart():
            for payl in em.get_payload():
                extract_attachment(payl)
        else:
            extract_attachment(em)

if __name__=='__main__':
    if len(sys.argv) < 2 or len(sys.argv) > 3:
        print "usage: %s <mbox_file> [directory]" % sys.argv[0]
        sys.exit(1)

    filename = sys.argv[1]
    directory = os.path.curdir

    if not os.path.exists(filename):
        print "file doesn't exist:", filename
        sys.exit(1)

    if len(sys.argv) == 3:
        directory = sys.argv[2]
        if not os.path.exists(directory) or not os.path.isdir(directory):
            print "Directory doesn't exist:", directory
            sys.exit(1)

    main(filename, directory)

    print "total attachments extracted:", attachments
    print "total attachments skipped:", skipped
