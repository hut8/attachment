#!/usr/bin/env python3
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
# The use of a Content-Type of multipart in a body part within another
# multipart entity is explicitly allowed. In such cases, for obvious reasons,
# care must be taken to ensure that each nested multipart entity must use a
# different boundary delimiter. See Appendix C for an example of nested
# multipart entities.
# The use of the multipart Content-Type with only a single body part may be
# useful in certain contexts, and is explicitly permitted.
# The only mandatory parameter for the multipart Content-Type is the boundary
# parameter, which consists of 1 to 70 characters from a set of characters
# known to be very robust through email gateways, and NOT ending with white
# space. (If a boundary appears to end with white space, the white space must
# be presumed to have been added by a gateway, and should be deleted.) It is
# formally specified by the following BNF

# Related RFCs: 2047, 2044, 1522

from __future__ import print_function

__author__ = "Pablo Castellano <pablo@anche.no>"
__author__ = ', '.join([__author__, "Liam Bowen <liambowen@gmail.com>"])
__license__ = "GNU GPLv3+"
__version__ = 1.4
__date__ = "2016-03-09"

import base64
import binascii
import email
import mailbox
import os
import sys
try:
    from tqdm import tqdm
except ImportError:
    tqdm = list
    print("progress bar library not found")
    print( "run: pip install tqdm")

BLACKLIST = ('signature.asc', 'message-footer.txt', 'smime.p7s')

class ExtractionError(Exception):
    pass

def extract_attachment(msg, destination):
    if msg.is_multipart():
        print(destination)
        raise ExtractionError("tried to extract from multipart")

    attachment_data = msg.get_payload(decode=True)

    orig_destination = destination
    n = 1
    while os.path.exists(destination):
        # TODO: Detect if it's just a duplicate
        destination = orig_destination + "." + str(n)
        n += 1

    fp = None
    try:
        with open(destination, "wb") as sink:
            sink.write(attachment_data)
    except IOError as e:
        print("io error while saving attachment: %s" % str(e))

def process_message(msg, directory):
    for part in msg.walk():
        if part.get_content_disposition() == 'attachment':
            filename = part.get_filename()
            if filename:
                print(filename)
                destination = os.path.join(directory, filename)
                extract_attachment(part, destination)
            else:
                print("found message with nameless attachment: %s" % msg['subject'])

def main():
    if len(sys.argv) < 2 or len(sys.argv) > 3:
        print("usage: %s <mbox_file> [directory]" % sys.argv[0])
        sys.exit(1)

    filename = sys.argv[1]
    directory = os.path.curdir

    if not os.path.exists(filename):
        print("file doesn't exist:", filename)
        sys.exit(1)

    if len(sys.argv) == 3:
        directory = sys.argv[2]
        if not os.path.exists(directory) or not os.path.isdir(directory):
            print("Directory doesn't exist:", directory)
            sys.exit(1)

    box = mailbox.mbox(filename)
    print("counting messages... ")
    message_count = len(box)
    print("done (found %s)" % message_count)

    for i in tqdm(range(message_count), ascii=True):
        msg = box.get_message(i)
        process_message(msg, directory)

if __name__ == '__main__':
    main()
