#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import print_function

__author__ = "Pablo Castellano <pablo@anche.no>"
__author__ = ', '.join([__author__, "Liam Bowen <liambowen@gmail.com>"])
__license__ = "GNU GPLv3+"
__version__ = 1.4
__date__ = "2016-03-09"

import email
import mailbox
import os
import sys
import logging
import fnmatch
import hashlib
try:
    from tqdm import tqdm
except ImportError:
    print("progress bar library not found")
    print( "run: pip install tqdm")
    exit(1)

BLACKLIST = set(['signature.asc', 'message-footer.txt', 'smime.p7s'])

class ExtractionError(Exception):
    pass

def extract_attachment(msg, destination):
    if msg.is_multipart():
        logging.error("tried to extract from multipart: %s" % destination)
        return

    attachment_data = msg.get_payload(decode=True)

    orig_destination = destination
    n = 1
    while os.path.exists(destination):
        destination = orig_destination + "." + str(n)
        n += 1

    fp = None
    try:
        with open(destination, "wb") as sink:
            sink.write(attachment_data)
    except IOError as e:
        logging.error("io error while saving attachment: %s" % str(e))

def wanted(filename):
    if filename in BLACKLIST:
        return False
    for ext in ['*.doc', '*.docx', '*.odt', '*.pdf', '*.rtf']:
        if fnmatch.fnmatch(filename, ext):
            return True
    return False

def process_message(msg, directory):
    for part in msg.walk():
        if part.get_content_disposition() == 'attachment':
            filename = part.get_filename()
            if filename and wanted(filename):
                logging.debug("extract filename: %s" % filename)
                destination = os.path.join(directory, filename)
                extract_attachment(part, destination)
            elif filename:
                logging.debug("found message with nameless attachment: %s" % msg['subject'])

def main():
    if len(sys.argv) < 2 or len(sys.argv) > 3:
        print("usage: %s <mbox_file> [directory]" % sys.argv[0])
        sys.exit(1)

    filename = sys.argv[1]
    directory = os.path.curdir

    logging.basicConfig(
        filename='attachment-%s.log' % os.path.basename(filename),
        level=logging.DEBUG)

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
