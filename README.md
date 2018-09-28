# IMAP Copy/SYNC
Copy emails and folders from an IMAP account to another one even in deifferent servers.

Creates missing folders and skips existing messages (using message-id).

Emails copied with their accurate timestamp so they don't appear as new(today) in the destination

Batch file included(startSyncProcess) with example CSV to copy/sync multiple mailboxes

use sync.py and remember to install pytz in python 3.X

Source IMAP is always accessed READ-ONLY.

```
Usage: imapcp.py <user>:<password>:<host>:<port> <user>:<password>:<host>:<port>

Options:
  --version             show program's version number and exit
  -h, --help            show this help message and exit
  -e EXCLUDE, --exclude=EXCLUDE
                        Exclude folders matching pattern (can be specified
                        multiple times)
  -f FOLDER, --folder=FOLDER
                        Only copy a single folder (use from:to to specify a
                        different destinatin name)
  -s, --simulate        Do not perform any task
  --from=FR             Only copy messages older than this date (inclusive)
  --to=TO               Only copy messages newer than this date (inclusive)
  --zone="timezone"     Timezone for source emails as in pytz lib
  ```

  If you are looking for a list of pytz timezones cheout this [gist](https://gist.github.com/heyalexej/8bf688fd67d7199be4a1682b3eec7568)
  
  c# project is hardly upto the task but soon :)

  forked and updated from [GTOZZI](https://github.com/gtozzi/imapcp)
  
