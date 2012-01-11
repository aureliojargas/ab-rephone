#!/usr/bin/env python2.6
# coding: utf-8
#
# ab-rephone.py
# 2011-06-20 by Aurelio Jargas
# Tested in Mac OS X 10.6.7
# License: MIT
#
# This script performs a regex-powered batch search/replace in the phone
# numbers of your contacts in Apple's Address Book app for Mac OS X. Useful
# for adding/removing local area codes, prefixes and general formatting.
# See CONFIG for details and run this script in Terminal.
#
# When running in normal mode (not dry-run), the script will ask you before
# making changes to each number, with the following options:
#
#     y    YES, make this change. This is the default action.
#     n    NO, don't make this change.
#     q    QUIT: save changes already made and quit.
#     a    ALL: make all changes without asking (be careful!)
#
# Note: Backup your contacts before anything else (File > Export...)
# Note: Quit Address Book before using this script to avoid conflicts.
#
# WARNING: DO NOT USE THIS SCRIPT IF YOU'RE NOT A REGEX NINJA.
#

import AddressBook
import sys
import re


#############################################################################
### CONFIG
#############################################################################
#
# By default all contacts are scanned.
# You can restrict the scan to a specific group with this setting.
#
group_name = ''    # Default=''

# Here are the regexes that match and replace the phone numbers.
# Use as many search/replace tuples as you need.
# Regexes are applied in order.
#
# The following samples are for adding the 041 prefix, used by the
# TIM carrier in Brazil. It also adds the area code (47) if missing.
# Examples:
# 1234-5678       -> (041 47) 1234-5678
# (47) 1234-5678  -> (041 47) 1234-5678
#
patterns = [

        # NNNNNNNN -> NNNN-NNNN
        # Add hyphen
        ('^(\d{4})(\d{4})$', r'\1-\2'),

        # NNNN-NNNN -> (47) NNNN-NNNN
        # Numbers with no area code. Add the (47) prefix.
        ('^(\d{4}-\d{4})$', r'(47) \1'),

        # (0AA) NNNN-NNNN -> (AA) NNNN-NNNN
        # Remove the leading zero in malformed phone (no carrier)
        ('^\(0(\d{2})\) (\d{4}-\d{4})$', r'(\1) \2'),

        # (AA) NNNN-NNNN -> (041 AA) NNNN-NNNN
        # Numbers with area code, but no carrier code (041)
        ('^\((\d{2})\) (\d{4}-\d{4})$', r'(041 \1) \2'),

        ### Some useful samples you may use:
        # 
        # # 0800NNNNNN   -> 0800-NNN-NNN
        # # 0800NNNNNNN  -> 0800-NNN-NNNN
        # # 0800NNNNNNNN -> 0800-NNNN-NNNN
        # # Format 0800 and 0300 special numbers
        # ('^(0[38]00)(\d{3})(\d{3})$', r'\1-\2-\3'),
        # ('^(0[38]00)(\d{3})(\d{4})$', r'\1-\2-\3'),
        # ('^(0[38]00)(\d{4})(\d{4})$', r'\1-\2-\3'),
        # 
        # # Change carrier, from Claro (21) to TIM (41)
        # # (021 AA) NNNN-NNNN -> (041 AA) NNNN-NNNN
        # ('^\(021 ', r'(041 '),
        # 
        # # Old numbers with 7 digits, add prefix 3
        # # NNNNNNN  -> 3NNN-NNNN
        # # NNN-NNNN -> 3NNN-NNNN
        # # (AA) NNN-NNNN -> (AA) 3NNN-NNNN
        # ('^(\d{3})(\d{4})$', r'3\1-\2'),
        # ('^(\d{3}-\d{4})$', r'3\1'),
        # ('^(\(\d{2}\)) (\d{3}-\d{4})$', r'\1 3\2'),
        # 
        # # International numbers (Brazil)
        # # +55AANNNNNNNN -> (041 AA) NNNN-NNNN
        # ('^\+55(\d{2})(\d{4})(\d{4})$', r'(041 \1) \2-\3'),
        #
        # # Add hyphen to the telephone number (by Michel de Almeida)
        # # (NNN NN) NNNNNNNN -> (NNN NN) NNNN-NNNN
        # ('^(\(\d{3} \d{2}\)) (\d{4})(\d{4})$', r'\1 \2-\3'),
]

# DRY RUN
# This mode shows the list of changes, but don't apply them.
# Your contacts won't be changed, it's a safe execution for testing.
# Test and fine tune your regexes. When they're OK, disable this mode.
#
dry_run = 1    # Default=1


# Some misc options you may never need to use.
max_people = 0  # Limit the number of people. Default=0 (OFF)
max_phones = 0  # Limit the number of phones. Default=0 (OFF)
show_unchanged = 1  # Show numbers that won't be changed? Default=1
#
#############################################################################
### END OF CONFIG
#############################################################################


# Do not change here
should_quit = 0
ab_has_changed = 0
change_all = 0

if dry_run:
        change_all = 1

def getPhones(person):
        """
        Returns a tuple with the phones record (ABMultiValueCoreDataWrapper)
        and a list with all the person's phones, each in the following format:
        [phone_uid, phone_number]

        Note: I'm saving UID not INDEX, because:

        http://developer.apple.com/library/mac/#documentation/UserExperience/Conceptual/AddressBook/Tasks/AccessingData.html
        You use the numeric index to access items in a multivalue list, but
        these indices may change as the user adds and removes values. If you
        want to save a reference to a specific value, use the unique
        identifier, which is guaranteed not to change.
        """
        phones = []
        phones_record = person.valueForProperty_('Phone')

        for index in range(0, phones_record.count()):
                value = phones_record.valueAtIndex_(index)
                uid = phones_record.identifierAtIndex_(index)
                phones.append([uid, value])
        return (phones_record, phones)


#############################################################################

# Dry run welcome message
# You may comment raw_input() when it becomes annoying (it will)
if dry_run:
        print("")
        print("The DRY RUN mode is turned ON.")
        print("")
        print("No changes will be made to your Address Book.")
        print("You can safely play with this script.")
        print("")
        print("See the CONFIG area at the script headers.")
        print("Now it's time to fine tune your regexes.")
        print("When they are perfect, disable DRY RUN.")
        print("")
        print("Now just press ENTER to continue...")
        raw_input()

# Show search summary before any processing
if group_name:
        where = "the group '%s'" % group_name
else:
        where = "all people in Address Book"

if max_people != 0 and max_phones != 0:
        limits = ", stoping with %s people or %s phones" % (max_people, max_phones)
elif max_people != 0:
        limits = ", stoping with %s people" % max_people
elif max_phones != 0:
        limits = ", stoping with %s phones" % max_phones
else:
        limits = ''

if group_name or limits:
        print('-' * 78)
        print("SUMMARY:\nWill search %s%s.\n" % (where, limits))


#############################################################################

# Start counters
person_count = 0
phone_count = 0

# Get the Address Book database
ab = AddressBook.ABAddressBook.sharedAddressBook()

# User specified a group?
# If so, limit the people to those inside this group
if group_name:

        # Search this group in AB
        people = None
        for group in ab.groups():
                props = group.allProperties()

                if props.get('GroupName').lower() == group_name.lower():

                        # Group found, now get the people in it
                        people = group.members()
                        break
        # Oops
        if not people:
                print("Error: Group '%s' not found or empty" % group_name)
                sys.exit(1)

# No group, let's get everybody
else:
        people = ab.people()

# Limit the number of people (if needed)
if max_people > 0:
        people = people[:max_people]


### Ok, we have the people list
### Now let's scan their phone numbers


for person in people:
        person_has_changed = 0

        # Get properties
        props = person.allProperties()

        # No phone, no deal
        if not 'Phone' in props:
                continue

        # Get person (or company) name
        name = " ".join([props.get('First', ''), props.get('Last', '')]).strip()
        if not name: name = props.get('Nickname')
        if not name: name = props.get('Organization')

        # Get phones: list of [uid, phone_nr] values
        phones_record, phones = getPhones(person)

        # Note: The returned phone record is read-only.
        # We must create a mutable copy, edit it and overwrite the original.
        # Thanks Johan Kool for the hint:
        # http://www.cocoabuilder.com/archive/cocoa/197150-abmutablemultivalue-error.html
        phones_record_mutable = phones_record.mutableCopy()

        # Search/replace phones
        for (uid, phone) in phones:

                old_ = phone
                new_ = phone

                # Limit the number of phones in output
                phone_count += 1
                if max_phones > 0 and phone_count > max_phones:
                        break

                # Apply the changes to this phone number
                for (pattern, replace) in patterns:
                        new_ = re.sub(pattern, replace, new_)

                # Anything has changed?
                if old_ != new_:
                        # YES
                        prompt = "%-20s %20s -> %20s   " % (name[:20], old_, new_)

                        if not change_all:
                                print prompt,  # to avoid line break
                                answer = raw_input('[Ynqa] ').lower()
                                # Note: raw_input is not Unicode-safe.

                                if answer.startswith('n'):  # No
                                        continue
                                elif answer.startswith('q'):  # Quit
                                        should_quit = 1
                                        break
                                elif answer.startswith('a'):  # All
                                        change_all = 1
                        else:
                                if dry_run:
                                        print(prompt + '[dry-run]')
                                else:
                                        print(prompt + 'YES')

                        # Will save the new phone number
                        if not dry_run:
                                index = phones_record_mutable.indexForIdentifier_(uid)
                                phones_record_mutable.replaceValueAtIndex_withValue_(index, new_)
                                person_has_changed = 1
                                ab_has_changed = 1

                else:
                        # NO, nothing to do.
                        if show_unchanged:
                                print("%-20s %20s" % (name[:20], old_))

        # Overwrite the person's phone record with the new data
        if person_has_changed and not dry_run:
                person.setValue_forProperty_(phones_record_mutable, 'Phone')

        # Quit is not a sys.exit() because we need to save the already made changes
        if should_quit:
                break

if ab_has_changed:
        print("\nSaving changes to the Address Book database... ")
        ab.save()
        print("Done.")


# Coding references (thank you!):
# http://www.programmish.com/?p=26
# https://gist.github.com/1029870
# https://github.com/prabhu/macutils/blob/master/addressbook.py
# http://www.mactech.com/articles/mactech/Vol.25/25.06/2506MacintheShell/index.html
# http://developer.apple.com/library/mac/#documentation/UserExperience/Conceptual/AddressBook/Tasks/AccessingData.html
# http://developer.apple.com/library/mac/#documentation/UserExperience/Reference/AddressBook/Classes/ABGroup_Class/Reference/Reference.html
# http://developer.apple.com/library/mac/#documentation/UserExperience/Reference/AddressBook/Classes/ABPerson_Class/Reference/Reference.html
# http://developer.apple.com/library/mac/documentation/UserExperience/Reference/AddressBook/Classes/ABAddressBook_Class/ABAddressBook_Class.pdf
# http://developer.apple.com/library/mac/#documentation/UserExperience/Reference/AddressBook/Classes/ABMutableMultiValue_Class/Reference/Reference.html
