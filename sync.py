import imaplib
import sys
import re
import pprint
import email
import datetime
import pytz

from optparse import OptionParser
from imaputil import ImapUtil

class main(ImapUtil):

    NAME = 'Sync'
    VERSION = '0.1'

    def run(self):

        # Init pretty printer
        pp = pprint.PrettyPrinter(indent = 2)
        #zone=pytz.timezone("Africa/Nairobi")
        log= open(r"C:\Users\Administrator\Desktop\IMAP Sync\log.txt","a+")

        # Read command line
        usage = "%prog <user>:<password>:<host>:<port> <user>:<password>:<host>:<port>"
        parser = OptionParser(usage=usage, version=self.NAME + ' ' + self.VERSION)
        parser.add_option("-e", "--exclude", dest="exclude", action='append',
            help="Exclude folders matching pattern (can be specified multiple times)")
        parser.add_option("-f", "--folder", dest="folder",
            help="Only copy a single folder (use from:to to specify a different destinatin name)")
        parser.add_option("-s", "--simulate", dest="simulate", action='store_true',
            help="Do not perform any task")
        parser.add_option("-t", "--trim", dest="trim", action='store_true',
            help="Trim folder names")
        parser.add_option("-k", "--skel", dest="skel", action='store_true',
            help="Only copy folder structure")
        parser.add_option("--from", dest="fr",
            help="Only copy messages older than this date (inclusive)")
        parser.add_option("--to", dest="to",
            help="Only copy messages newer than this date (inclusive)")
        parser.add_option("--zone",dest="zone",help="Timezone of source server clock")

        (options, args) = parser.parse_args()

        # Parse exclude list
        excludes = []
        if options.exclude:
            for e in options.exclude:
                excludes.append(re.compile(e.encode()))

        # Parse from/to dates
        zone=pytz.timezone("Africa/Nairobi");
        if options.zone:
            zone=pytz.timezone(options.zone)
        fr = None
        if options.fr:
            fr = datetime.datetime(*[int(i) for i in options.fr.split('-')])
            fr=fr.replace(tzinfo=zone)           
        if fr:
            print("*************************************************")
        to = None
        if options.to:
            to = datetime.datetime(*[int(i) for i in options.to.split('-')])
            to=to.replace(tzinfo=zone)
        if to:
            print("Only copying messages older than %s (included)" % to)

        # Parse single folder
        folder = options.folder.split(':') if options.folder else None
        if folder and len(folder) < 2:
            folder.append(folder[0])
        if folder:
            print("Only copying folder %s to folder %s" % (folder[0], folder[1]))

        # Parse mandatory arguments
        if len(args) < 2:
            parser.error("invalid number of arguments")
        src = args[0].split(':')
        src = {
            'user': src[0],
            'pass': src[1],
            'host': src[2] if len(src) > 2 else 'localhost',
            'port': int(src[3]) if len(src) > 3 else 143,
        }
        dst = args[1].split(':')
        dst = {
            'user': dst[0],
            'pass': dst[1],
            'host': dst[2] if len(dst) > 2 else 'localhost',
            'port': int(dst[3]) if len(dst) > 3 else 143,
        }

        # Make connections and authenticate
        if src['port'] == 993:
            srcconn = imaplib.IMAP4_SSL(src['host'], src['port'])
        else:
            srcconn = imaplib.IMAP4(src['host'], src['port'])
        try:
            srcconn.login(src['user'], src['pass'])
        except imaplib.IMAP4.error:
            print("Login for user ",src['user']," failed at source server")
            dv="Login for user "+str(src['user'])+" failed at source server"+"\r\n"
            log.write(dv)
            log.close()
            sys.exit(0)
        srctype, srcdescr = self.getServerType(srcconn)
        #print("Source server type is", srcdescr)

        if dst['port'] == 993:
            dstconn = imaplib.IMAP4_SSL(dst['host'], dst['port'])
        else:
            dstconn = imaplib.IMAP4(dst['host'], dst['port'])
        try:
            dstconn.login(dst['user'], dst['pass'])
        except imaplib.IMAP4.error:
            print("Login for user ",dst['user']," failed at destination server")
            dv="Login for user "+str(dst['user'])+" failed at destination server"+"\r\n"
            log.write(dv)
            log.close()
            sys.exit(0)
        dsttype, dstdescr = self.getServerType(dstconn)      

        print("Syncing email ",dst['user'])
        #log.write("\r\n*************************************************\r\n")
        #dv="Syncing email " + str(dst['user'])+"\r\n"
        #log.write(dv)

        #print("Source folders:")
        srcfolders = self.listMailboxes(srcconn)
        #print("Found", len(srcfolders), "folders in  PostFIX")        
        dstfolders = self.listMailboxes(dstconn)
        #print("Found", len(dstfolders), "folders in Exchange")

        # Syncing every source folder
        
        started=datetime.datetime.now()
        msgcopied=0
        msgskipped=0
        for f in srcfolders:
            srcfolder = '"{}"'.format(f.name.decode("utf-8"))
            dstfolder = '"{}"'.format(f.getPathBytes(dsttype, trim=options.trim).decode("utf-8"))

            # Create dst folder when missing
            dstconn.create(dstfolder)


            # Select source mailbox readonly
            res, data = srcconn.select(srcfolder, True)
            if res == 'NO' and srctype == 'exchange' and 'special mailbox' in data[0]:
                print("Skipping special Microsoft Exchange Mailbox", srcfolder)
                continue
            assert res == 'OK', (res, data)
            res, data = dstconn.select(dstfolder, False)
            if res == 'OK':
                pass
            elif res == 'NO':
                print('Error selecting folder: {}, trying to create it'.format(str(data)))
                # Create and try again
                res, data = dstconn.create(dstfolder)
                if res != 'OK':
                    raise RuntimeError('Error creating mailbox "{}": {}'.format(dstfolder.decode(), str(data)))
                res, data = dstconn.select(dstfolder, False)
                assert res == 'OK', (res, data)
            else:
                assert False, (res, data)

            # Stop here if only copying skeleton
            if options.skel:
                print("Skipping message copy")
                continue

            # Fetch all destination messages imap IDS
            dstids = self.listMessages(dstconn)
            #print("Found", len(dstids), "messages in destination folder")

            # Fetch destination messages ID
            #print("Acquiring destination message IDs...", end='', flush=True)
            dstmexids = []
            for idx, did in enumerate(dstids):
                if idx % 100 == 0:
                    print('.', end='', flush=True)
                dstmexids.append(self.getMessageId(dstconn, did))
            #print(len(dstmexids), "message IDs acquired.")

            # Fetch all source messages imap IDS
            srcids = self.listMessages(srcconn)
            #print("Found", len(srcids), "messages in source folder")
            msgcopied=0
            msgfound=len(srcids)            
            date=datetime.datetime(2018,9,26,0,0,0,0,zone)

            # Sync data
            for sid in srcids:
                # Check for date filter
                if fr or to:
                    h = self.getHeaders(srcconn, sid)
                    if 'date' not in h:
                        continue
                    d = email.utils.parsedate(h['date'])
                    if not d:
                        continue   
                    date = datetime.datetime(d[0], d[1], d[2], d[3], d[4], d[5], 0, zone)
                    if fr and date < fr:
                        continue
                    if to and date > to:
                        continue
                # Get message id
                mid = self.getMessageId(srcconn, sid)
                if not mid in dstmexids:
                    # Message not found, syncing it
                    #print("Copying message", mid)
                    if not options.simulate:
                        mex = self.getMessage(srcconn, sid)
                        try:
                            res=dstconn.append(dstfolder, None, date, mex)
                            msgcopied+=1
                        except Exception as ex:
                            print(ex)
                            if dst['port'] == 993:
                                dstconn = imaplib.IMAP4_SSL(dst['host'], dst['port'])
                            else:
                                dstconn = imaplib.IMAP4(dst['host'], dst['port'])
                            dstconn.login(dst['user'], dst['pass'])
                            dstconn.select(dstfolder, False)
                            msgskipped+=1
                            #print("Skipping message", mid," in folder ",dstfolder)
                            #res=dstconn.append(dstfolder, None, None, mex)
                """else:
                    print("Skipping message", mid)"""

        # Logout
        srcconn.logout()
        dstconn.logout()
        ended=datetime.datetime.now()
        print("Syncing ended after ", (ended-started).total_seconds() / 60.0," mins ",msgskipped," >>> ", msgcopied,"/",msgfound)
        #log.write("Syncing ended after "+ str((ended-started).total_seconds() / 60.0)+" mins "+str(msgskipped)+" >>> "+ str(msgcopied)+"/"+str(msgfound))
        log.close()

        if options.simulate:
            print("Simulated run, no action taken")


if __name__ == '__main__':
    app = main()
    app.run()
    sys.exit(0)
