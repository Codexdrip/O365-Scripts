#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os, sys
import win32com.client
from datetime import datetime
from datetime import timedelta, time
import pywintypes
import time

#=============================================Funct/Variables/Const=======================================
# Microsoft Outlook Constants
# http://msdn.microsoft.com/en-us/library/aa219371(office.11).aspx
# or you can use the following command to generate
# c:\python26\python.exe c:\Python26\lib\site-packages\win32com\client\makepy.py -d
# After generated, you can use win32com.client.constants.olFolderSentMail
# http://code.activestate.com/recipes/496683-converting-ole-datetime-values-into-python-datetim/


def loadOutlook():
    app = win32com.client.Dispatch( "Outlook.Application" ).GetNamespace( "MAPI" )
    return app

olFolderDeletedItems=3
olFolderSentMail=5
olFolderInbox=6
OLE_TIME_ZERO = datetime(1899, 12, 30, 0, 0, 0)
mapi = loadOutlook()

default_folders = [
        #mapi.GetDefaultFolder(olFolderSentMail),
        #mapi.GetDefaultFolder(olFolderDeletedItems)
        mapi.GetDefaultFolder(olFolderInbox)
]
#====================================================== Main Program ====================================================# 
if __name__ == '__main__':
    for folder in default_folders:
        print "[!] Processing %s" % folder.Name
        print 'How many days do you want to go back?'
        num_of_days = int(raw_input('>> '))
        mark2delete=[]
        #If you use makepy.py, you have to use the following codes imapitead of "for item in folder.Items"
        #for i in range(1,folder.Items.Count+1):
        #    item = folder.Items[i]
        try:
            for i in range(1,folder.Items.Count+1):
                inbox = folder.Items
                msg = inbox[i]
                recv_time = datetime.strptime(str(msg.CreationTime), "%m/%d/%y %H:%M:%S")
                past30days=datetime.now()-timedelta(days=num_of_days) # the date 10 days ago
                try:
                    if recv_time < past30days: # ex: if 2017-01 < 2018-01 then delete the message                   
                        #os.system('cls')
                        mark2delete.append(msg)
                        #print 'Number of items added: {}'.format(len(mark2delete))
                        print '{0} < {1} = True'.format(recv_time, past30days)
                    else:
                        print '[!] Found last item...'
                        time.sleep(3)
                        break
                except AttributeError:
                    mark2delete.append(msg)
        except IndexError:
            print '[!] caught'
        #==================================================Second Act========================================================
            
        if len(mark2delete)>0:
            os.system('cls')
            for x, item in enumerate(mark2delete):
                try:
                    if x%3 == 0:
                        print '\nSubject: {0} |'.format(item),
                    else:
                        print 'Subject: {0} |'.format(item),
                except AttributeError:
                    print 'Subject: unknown |',
        
            #===============================================Finally==========================================================
            cont = raw_input('\n\n[!] Im about to delete {} items! Continue?'.format(len(mark2delete)))
            if cont == 'Y':
                for item in mark2delete:
                    print "[-] Removing: %s" % item 
                    item.Delete()
            else:
                print '[-] Closing program...'
            print '[!] Done!!!'
        else:
            print "[!] No matched mail."