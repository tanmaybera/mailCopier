import sys
import email
import time
import imaplib
import datetime


# """----------------- account credentials ------------"""
imapserver = "imap-mail.outlook.com"                  # I used outlook imap server
username = "abc@live.com"                             # example : "abc@live.com"   [Your user name]
password = "qwe123"                                   # example : "123qwe"         [Your password]

# ****** BE CARE FULL FOR BELOW variable EDIT , Should correct
# ****** BE CARE FULL FOR BELOW variable EDIT , Should correct

Source_mailfolder = "Inbox"            #example: "Inbox/love"
Dest_mailfolder = "Inbox/love"         #example: "Inbox/Other"

#^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

print("\n-----------------------------------------------------------------------------")
print("--*                    Mail Copier Process                                *--")
print("-----------------------------------------------------------------------------")


if (Source_mailfolder == Dest_mailfolder ):
    print("ERROR : Source MAIL Folder and Destination Mail Folder CANT NOT BE SAME.")
    print("Process :  EXIT")
    sys.exit(1)
elif (len(Source_mailfolder)== 0 or len(Dest_mailfolder)==0):
    print("ERROR : Source MAIL Folder / Destination Mail Folder CANT NOT BE BLANK.")
    print("Process :  EXIT")
    sys.exit(1)
else:
    print("\nParameter:")
    print("Source Mail Folder     : "+str(Source_mailfolder))
    print("Destination Mail Folder: "+str(Dest_mailfolder))
    pass


print("\n* Please note that START DATE is OLDER DATE than END DATE *")
startdate = input("\nPlease enter START DATE (MM/DD/YYYY) :")
enddate = input("\nPlease enter END DATE (MM/DD/YYYY) :")

try:
    sdate = datetime.datetime.strptime(startdate, '%m/%d/%Y')
except:
    print("Error in START DATE format. Pls check")
    sys.exit(1)

try:
    edate = datetime.datetime.strptime(enddate, '%m/%d/%Y')
except:
    print("\n")
    print("Error in END DATE format. Pls check")
    print("\n")
    print("Process :  EXIT")
    sys.exit(1)


if sdate < edate:
    pass
else:
    print("\nPLEASE READ CAREFULLY : * Please note that START DATE is OLDER DATE than END DATE *")

imap = imaplib.IMAP4_SSL(imapserver)
imap.login(username, password)
status, messages = imap.select(Source_mailfolder, readonly = False)
messages = int(messages[0])
print("\nPLEASE READ CAREFULLY : \n")
print("    1. Will COPY mails FROM ["+str(Source_mailfolder)+ "] TO ["+str(Dest_mailfolder)+"].")
print("    2. Will COPY mails BETWEEN ["+str(enddate)+ "]  AND  ["+str(startdate)+"].")
print("\nWARN: You have TOTAL mails in ["+Source_mailfolder+"] : "+str(messages))
ll = input("\nDo you want to [START NOW]?.  Please confirm (Y/N) : ")

if ll.upper() == "Y":
    print("\nYou have entered Option : [" + str("Y")+"].")
elif ll.upper() == "N":
    print("\nYou have entered Option : [" + str("N") + "].")
    print ("\nProcess :  EXIT")
    sys.exit(1)
else:
    print("\nYou have entered WRONG OPTIONs : "+str(ll))
    print("\nProcess :  EXIT")
    sys.exit(1)


print("\nProcess has been Started ....")
st = time.time()

date_generated = [sdate + datetime.timedelta(days=x) for x in range(0, (edate-sdate).days+1)]
total_mail = 0
mail_moved = 0
mail_error = 0

for i in date_generated:
    no_mail = 0
    cp_mail = 0

    try:
        resp, data = imap.search(None, "(ON {0})".format(i.strftime("%d-%b-%Y")))
    except:
        print("DATE WISE STAT : ["+i.strftime("%d-%b-%Y")+"]. Total Mail : [0]. Successfully Copied : [0]. Error Copy: [0]. #")
        continue

    midl = "".join([str(i.decode()) for i in (list(data))]).split()
    total_mail = total_mail + len(midl)


    if len(midl) == 0:
        print("DATE WISE STAT : ["+i.strftime("%d-%b-%Y")+"]. Total Mail : [0]. Successfully Copied : [0]. Error Copy: [0]. *")
        continue

    #print("STAT : TOTAL MAILS for ["+i.strftime("%d-%b-%Y")+"] data is : "+str(len(midl)))

    for j in midl:
        id = str(j).strip().encode()
        resp, data1 = imap.fetch(id, "(UID)")

        uid_str = "".join(str(i) for i in data1)
        uid_tmp_s = uid_str.find("UID")
        uid_tmp_e = uid_str.find(")")

        UID = (''.join(filter(str.isdigit, uid_str[uid_tmp_s:uid_tmp_e])))
        try:
            result = imap.uid('COPY', UID, Dest_mailfolder)
            cp_mail = cp_mail + 1
            mail_moved = mail_moved + 1
        except:
            no_mail = no_mail + 1
            mail_error = mail_error + 1

    print("DATE WISE STAT : ["+i.strftime("%d-%b-%Y")+"]. Total Mail : ["+str(len(midl))+"]. Successfully Copied : ["+str(cp_mail)+"]. Error Copy: ["+str(no_mail)+"].")


print("\nProcess has been Completed.")
print ("\n---* REPORT *---")
print ("\n* FROM ["+Dest_mailfolder+"] TO ["+Source_mailfolder+"] * || * BETWEEN ["+str(enddate)+ "]  AND  ["+str(startdate)+"] *")
print ("\n* Total Mail : ["+str(total_mail)+"]")
print ("\n* Total Copy : ["+str(mail_moved)+"]")
print ("\n* Total ERR  : ["+str(mail_error)+"]")

imap.close()
imap.logout()
print('\nTotal Process Time : ', time.time()-st, 'seconds.\n')
print ("\n---# REPORT #---\n\n")
