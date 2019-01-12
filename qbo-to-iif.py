#!/usr/bin/python

###
###  Script to convert a Quickbooks WebConnect QBO file to IIF format
###
###  Takes the QBO file to convert, and the Quickbooks account that
###  the IIF file will be imported into, as parameters:
###
###  Usage: qbo-to-iif.py to-convert.QBO "QB Account to Import Into"
###
###  IMPORTANT NOTE: The Quickbooks account given as a parameter must
###  match the name in Quickbooks (obviously), but if it's a
###  sub-account, it needs to have the colon separator. For example,
###  if you have a Foo Bank account which is a sub-account under Bank
###  Accounts, you'll want to give "Bank Accounts:Foo Bank"
###  
###  This may be necessary if your copy of Quickbooks is over 3 years
###  old and refuses to import WebConnect files (.QBO files) any more
###  - as a way for Intuit to force you to upgrade every 3 years.
###
###  The error message Quickbooks gives when you try to import is:
###  "QuickBooks is unable to verify the Financial Institution
###  information for this download. Please try again later."
###
###  Note that this message can also arise if your Quickbooks version
###  isn't over 3 years old, but if your bank refused to pay Intuit
###  the licensing fee to be able to let its customers download in
###  WebConnect format.
###  
###  In that case, you may be able to edit the QBO file (with a text
###  editor like Notepad or TextEdit) and change the <ORG>, <FI>, and
###  <INTU.BID> fields to the name/numbers of a bank that did pay the
###  Intuit licensing fee.
###
###  Copyright (C) 1998-2012 by the Free Software Foundation, Inc.
###
###  This program is free software; you can redistribute it and/or
###  modify it under the terms of the GNU General Public License
###  as published by the Free Software Foundation; either version 2
###  of the License, or (at your option) any later version.
###
###  This program is distributed in the hope that it will be useful,
###  but WITHOUT ANY WARRANTY; without even the implied warranty of
###  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
###  GNU General Public License for more details.
###
###  You should have received a copy of the GNU General Public License
###  along with this program; if not, write to the Free Software
###  Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA
###  02110-1301, USA. http://www.fsf.org/licensing/licenses/gpl.html
###
###  Author: Anthony R. Thompson
###  Contact: put @ between art and sigilservices.com
###  Created: April 2012
###

import os, sys, re, datetime
from ofxparse import OfxParser

# Initial settings - may need to change these to match your QB setup
payeefixfile = './payee-fixes.txt'
payeeacctfile = './payees-to-accounts.txt'
txtypesmapfile = './qbo-iif-txtypes.txt'
defaultcat = 'Miscellaneous - ASK ACCOUNTANT'

if (len(sys.argv) < 3):
    print 'Usage: ' + sys.argv[0] + " file.qbo 'QB Import Account Name'"
    sys.exit(1)
else:
    qbofile = sys.argv[1]
    impacct = sys.argv[2]
iiffile = re.sub('\.qbo$' + '(?i)', '.iif', qbofile)  # ignorecase

# Check for existence of configuration files before loading them below
class MissingConfigFileException(Exception): pass
for cf in [payeefixfile, payeeacctfile, txtypesmapfile]:
    if not os.path.exists(cf):
        raise MissingConfigFileException("Missing config file %s" % (cf))

# Regexp for getting non-comment, non-blank lines from prop files
noncomre = re.compile(r'^[^#]\S')

# Read in file that maps QBO transaction types to IIF ones
qbo_iif_txtypes = {}
for line in map(lambda x: x[:len(x)-1],   # eliminate newlines at end
                filter(lambda x: noncomre.search(x),
                       (open(txtypesmapfile, 'r')).readlines())):
    qbotxtype, iiftxtype = line.split("\t", 1)
    qbo_iif_txtypes[qbotxtype] = iiftxtype
class UnknownTxTypeException(Exception): pass

# Read in payee name fixes file
payeefixes = []
payeesubs = {}
asteriskre = re.compile('\*')
for line in map(lambda x: x[:len(x)-1],   # eliminate newlines at end
                filter(lambda x: noncomre.search(x),
                       (open(payeefixfile, 'r')).readlines())):
    orig, fixed = line.split("\t", 1)
    orig = asteriskre.sub('\*', orig)
    newre = re.compile('^' + orig + '(?i)')
    payeefixes.append(newre)
    payeesubs[newre] = fixed

# Regexp hack for B of A debit card transactions (remove CHECKCARD 0516)
checkcardre = re.compile('^CHECKCARD \d{4} ')
paypalre = re.compile('^PAYPAL \*')
paypaldesre = re.compile('^PAYPAL DES:INST XFER ID:')
ppre = re.compile('^PP\*')
squarere = re.compile ('^SQ \*')
tstre = re.compile ('^TST\* ')

def fix_payee(payee):
    payee = checkcardre.sub('', payee)
    payee = paypaldesre.sub('', payee)
    payee = paypalre.sub('', payee)
    payee = ppre.sub('', payee)
    payee = squarere.sub('', payee)
    payee = tstre.sub('', payee)
    for matchre in payeefixes:
        if matchre.search(payee): return payeesubs[matchre]
    return payee

# Read in payee -> quickbooks account mapping file
payeeaccts = {}
for line in map(lambda x: x[:len(x)-1],   # eliminate newlines at end
                filter(lambda x: noncomre.search(x),
                       (open(payeeacctfile, 'r')).readlines())):
    payee, acct = line.split("\t", 1)
    payeeaccts[payee] = acct

def acct_from_payee(payee):
    if payeeaccts.has_key(payee): return payeeaccts[payee]
    else: return defaultcat

# Use ofxparse library (yay) to read in QBO file
ofx = OfxParser.parse(file(qbofile))

# Write header for IIF file
out = open(iiffile, 'w')
out.write("!TRNS\tDATE\tTRNSTYPE\tDOCNUM\tNAME\tAMOUNT\tACCNT\tMEMO\t\n")
out.write( "!SPL\tDATE\tTRNSTYPE\tDOCNUM\tNAME\tAMOUNT\tACCNT\tMEMO\t\n")
out.write("!ENDTRNS\t\n")

# Write out each transaction
ampre = re.compile(r'&amp;')
for tx in ofx.account.statement.transactions:
    txdate = (tx.date).strftime("%m/%d/%y")
    if qbo_iif_txtypes.has_key(tx.type): txtype = qbo_iif_txtypes[tx.type]
    else:
        raise UnknownTxTypeException('Unknown transaction type ' +str(tx.type))
    txdocnum = re.sub('^0+', '', tx.checknum)
    txmemo = ampre.sub(r'&', tx.memo)
    txname = fix_payee(ampre.sub(r'&', tx.payee))
    txacct = acct_from_payee(txname)
    txamt = str(tx.amount)
    txrevamt = str(tx.amount * -1)

    out.write("TRNS\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t\n" % \
              (txdate, txtype, txdocnum, txname, txamt, impacct, txmemo))
    out.write("SPL\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t\n" % \
              (txdate, txtype, txdocnum, txname, txrevamt, txacct, txmemo))
    out.write("ENDTRNS\t\n")
    print "Processed %s / %s / %s / %s" % (txdate, txamt, txname, txacct)

out.close()
print "Wrote %s" % (iiffile)
