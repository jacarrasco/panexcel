#!/usr/bin/env python

from lxml import etree # Difficult to use but good XML parser.
import xlsxwriter # Creates an Excel Spreadsheet.
import argparse

class Spreadsheet(object):
    """Create a spreadsheet from the XML document."""
    def __init__(self):
        """Initial values to avoid issues when empty values"""
        self.name = None
        self.from_member = ""
        self.to_member = ""
        self.source = ""
        self.source_trans = ""
        self.destination = ""
        self.destination_trans = ""
        self.application = ""
        self.service = ""
        self.hipprofiles = ""
        self.action = None
        self.description = None
        self.logstart = None
        self.logend = None
        self.tag = ""
        self.profilesetting = ""
        self.disabled = "no" # Set to no since the PAN might return nothing for permit.
        self.expiration = None
        self.rulesection = ""
        self.ruletype = ""
        self.firewall= ""

        self.groupname= ""
        self.objectname = ""
        self.objectvalue = ""
        self.objectdescription = ""
        self.objectFQDN = "no"
        self.objectshared = "no"

    def writeRowHeaders(self):
        """Write the header row of the rule sheet."""
        titles = ["Rule Name", "From Zone", "To Zone", "Source", "Source Translation", "Destination", "Destination Translation", "Application", "Services", "Hip-Profile", "Action", "Description", "Log-start", "Log-end", "Tags", "Profile-settings", "Disabled", "Expiration", "Rule Section", "Rule Type", "Firewall"]
        i = 0
        for title in titles:
            worksheet.write(0, i, title, bold)
            i += 1

    def writeObjectHeaders(self):
        """Write the header row of the object sheet"""
        titles = ["Object Name", "Object Value", "Description", "FQDN", "Shared"]
        i = 0
        for title in titles:
            worksheet_objects.write(0, i, title, bold)
            i += 1

    def setName(self, name):
        """Populate the firewall rule description."""
        self.name = name

    def setFromMember(self, from_member):
        """Set firewall from zone."""
        if not self.from_member == "": # If there are multiple entries add a comma to separate.
            self.from_member += chr(10)
        self.from_member +=str(from_member) # Concatenate each entry.

    def setToMember(self, to_member):
        """Set firewall to zone."""
        if not self.to_member == "": # If there are multiple entries add a comma to separate.
            self.to_member += chr(10)
        self.to_member +=str(to_member) # Concatenate each entry.

    def setSource(self, source):
        """Set firewall from source."""
        if not self.source == "": # If there are multiple entries add a comma to separate.
            self.source += chr(10)
        self.source +=str(source) # Concatenate each entry.

    def setSourceTranslation(self, source_trans):
        """Set firewall from source Translation in NAT rules"""
        if not self.source_trans == "": # If there are multiple entries add a comma to separate.
            self.source_trans += chr(10)
        self.source_trans +=str(source_trans) # Concatenate each entry.

    def setDestination(self, destination):
        """Set firewall to destination."""
        if not self.destination == "": # If there are multiple entries add a comma to separate.
            self.destination += chr(10)
        self.destination +=str(destination) # Concatenate each entry.

    def setDestinationTranslation(self, destination_trans):
        """Set firewall from destination Translation for NAT rules."""
        if not self.destination_trans == "": # If there are multiple entries add a comma to separate.
            self.destination_trans += chr(10)
        self.destination_trans +=str(destination_trans) # Concatenate each entry.

    def setApplication(self, application):
        """Set firewall to application."""
        if not self.application == "": # If there are multiple entries add a comma to separate.
            self.application += chr(10)
        self.application +=str(application) # Concatenate each entry.

    def setServices(self, service):
        """Set firewall to Services."""
        if not self.service == "": # If there are multiple entries add a comma to separate.
            self.service += chr(10)
        self.service +=str(service) # Concatenate each entry.

    def setHipprofiles(self, hipprofiles):
        """Set HIP-Profile applied"""
        if not self.hipprofiles == "": # If there are multiple entries add a comma to separate.
            self.hipprofiles += chr(10)
        self.hipprofiles +=str(hipprofiles) # Concatenate each entry.

    def setAction(self, action):
        """Populate the firewall rule action."""
        self.action = action

    def setDescription(self, description):
        """Populate the firewall description."""
        self.description = description

    def setLogstart(self, logstart):
        """Set if rule has logstart."""
        self.logstart = logstart

    def setLogend(self, logend):
        """Set if rule has Logend."""
        self.logend = logend

    def setTag(self, tag):
        """Set firewall TAGS"""
        if not self.tag == "": # If there are multiple entries add a comma to separate.
            self.tag += chr(10)
        self.tag +=str(tag) # Concatenate each entry.

    def setProfilesetting(self, profilesetting):
        """Set Profile settings."""
        if not self.profilesetting == "": # If there are multiple entries add a comma to separate.
            self.profilesetting += chr(10)
        self.profilesetting +=str(profilesetting) # Concatenate each entry.

    def setDisabled(self, disabled):
        """Set if rule is disabled."""
        self.disabled = disabled

    def setExpiration(self, expiration):
        """Set if rule has expiration object."""
        self.expiration = expiration

    def setRulesection(self, rulesection):
        """Set rulesection."""
        self.rulesection = rulesection

    def setRuletype(self, ruletype):
        """Set ruletype."""
        self.ruletype = ruletype

    def setFirewall(self, firewall):
        """Set Firewall or VSYS."""
        self.firewall = firewall

    def setGroupname(self, groupname):
        """Set Groupname."""
        self.groupname = groupname

    def setObjectname(self, objectname):
        """Populate object name."""
        self.objectname = objectname

    def setObjectvalue(self, objectvalue):
        """Populate object value."""
        self.objectvalue = objectvalue

    def setObjectdescription(self, objectdescription):
        """Populate object description."""
        self.objectdescription = objectdescription

    def setObjectFQDN(self, objectFQDN):
        """Populate if object is FQDN."""
        self.objectFQDN = objectFQDN

    def setObjectShared(self, objectshared):
        """Populate if object is Shared."""
        self.objectshared = objectshared

    def writeRow(self, row):
        """Writes row to Excel workbook"""
        # Insert validation later
        worksheet.write(row, 0, self.name, dataformat)
        worksheet.write(row, 1, self.from_member, dataformat)
        worksheet.write(row, 2, self.to_member, dataformat)
        worksheet.write(row, 3, self.source, dataformat)
        worksheet.write(row, 4, self.source_trans, dataformat)
        worksheet.write(row, 5, self.destination, dataformat)
        worksheet.write(row, 6, self.destination_trans, dataformat)
        worksheet.write(row, 7, self.application, dataformat)
        worksheet.write(row, 8, self.service, dataformat)
        worksheet.write(row, 9, self.hipprofiles, dataformat)
        worksheet.write(row, 10, self.action, dataformat)
        worksheet.write(row, 11, self.description, dataformat)
        worksheet.write(row, 12, self.logstart, dataformat)
        worksheet.write(row, 13, self.logend, dataformat)
        worksheet.write(row, 14, self.tag, dataformat)
        worksheet.write(row, 15, self.profilesetting, dataformat)
        worksheet.write(row, 16, self.disabled, dataformat)
        worksheet.write(row, 17, self.expiration, dataformat)
        worksheet.write(row, 18, self.rulesection, dataformat)
        worksheet.write(row, 19, self.ruletype, dataformat)
        worksheet.write(row, 20, self.firewall, dataformat)

        print "Name: ", self.name
        print "From Zone: ", self.from_member
        print "To Zone: ", self.to_member
        print "Source: ", self.source
        print "Destination: ", self.destination
        print "Application: ", self.application
        print "Service", self.service
        print "Action: ", self.action
        print "Disabled: ", self.disabled
        print "Description: ", self.description
        print "Expiration: ", self.expiration
        print "\n"

    def writeObjectRow(self, row):
        """Writes row to Excel workbook"""
        # Insert validation later
        worksheet_objects.write(row, 0, self.objectname, dataformat)
        worksheet_objects.write(row, 1, self.objectvalue, dataformat)
        worksheet_objects.write(row, 2, self.objectdescription, dataformat)
        worksheet_objects.write(row, 3, self.objectFQDN, dataformat)
        worksheet_objects.write(row, 4, self.objectshared, dataformat)


    def newRow(self):
        """Prepares for new row by clearing variables in class"""
        excelobj.__init__()

    def newObjectRow(self):
        """Prepares for new row by clearing variables in class"""
        excelobjects.__init__()



def commandlineparser():
    """Select the proper arguments needed"""
    global args
    parser = argparse.ArgumentParser(description='Export rules and objects from Panorama Palo Alto Networks into an excel file')
    parser.add_argument('-f', '--firewall', required=False, help='Select a concrete Firewall Group or VSYS Name.')
    parser.add_argument('-v', '--virtualsystem', action="store_true", required=False, default=False, help='Enable if xml file has Virtual Systems instead of Device Groups')
#    parser.add_argument('-n', '--nat', required=False, help='Nat rules')
    parser.add_argument('-c', '--configfile', required=False, help='Introduce a concrete config file. By default config.xml is read')
    parser.add_argument('-r', '--rulename', required=False, help='Introduce a concrete rulename to obtain a report for just one rule')
    parser.add_argument('-o', '--objectname', required=False, help='Introduce a concrete object to obtain a report of all rules associated to that object')
    parser.add_argument('-e', '--excelname', required=False, help='Select a concrete output filename')
    args = parser.parse_args()

def getObjects(elementtree):
    """ Get all objects from the xml file"""
    objectrow=0
    for address in elementtree.iter("address"):
        #print address.tag
        for entry in address:
            objectrow +=1 
            excelobjects.setObjectname(objectname=entry.attrib.get("name"))
            #print entry.attrib.get("name")

            for netmask in entry.findall('ip-netmask'):
                excelobjects.setObjectvalue(netmask.text)

            for fqdn in entry.findall('fqdn'):
                excelobjects.setObjectvalue(fqdn.text)

            for description in entry.findall('description'):
                excelobjects.setObjectdescription(description.text)

            if address.getparent().tag == "shared":
                excelobjects.setObjectShared("yes")

            excelobjects.writeObjectRow(objectrow)
            excelobjects.newObjectRow()

def findbyobjectname(elementtree2, objectfind, objectrow):

    for address in elementtree2.iter("address"):
        for entry in address.findall(".//*[@name='%s']" %objectfind):
            objectrow+=1
            print entry.attrib.get("name")

            for netmask in entry.findall('ip-netmask'):
                excelobjects.setObjectvalue(netmask.text)

            for fqdn in entry.findall('fqdn'):
                excelobjects.setObjectvalue(fqdn.text)

            for description in entry.findall('description'):
                excelobjects.setObjectdescription(description.text)

            if entry.getparent().tag == "shared":
                excelobjects.setObjectShared("yes")

            excelobjects.writeObjectRow(objectrow)
            excelobjects.newObjectRow()
    return objectrow



if __name__ == '__main__':

    #Get command line arguments
    commandlineparser()

    row = 0 # Used to track which excel row we are on while parsing XML.
    objectrow=0
    rulesfound=[]
    allobjects = True

    if args.configfile == None:
        document = etree.parse('policy-best-practices.xml').getroot() # Parse the page the firewall returned as a string into the document object.
    else:
        document = etree.parse(args.configfile).getroot()

    if args.firewall == None:
        firewall_name=']'
    else:
        firewall_name="='"+args.firewall+"']"
        print firewall_name

    # if args.rulebase== None:
    #     rule_type='pre-rulebase'
    # else:
    #     rule_type=args.rulebase

    # if args.nat == None:
    #     security_or_nat='security'
    # else:
    #     security_or_nat=args.nat


    if args.virtualsystem== False:
        virtual_or_deviceg='device-group'
    else:
        virtual_or_deviceg='vsys'


    if args.objectname== None and args.rulename == None:
        rulename="]"
        rulesfound.append(rulename)
    if args.objectname != None and args.rulename == None:
        hostname=".//*[member='"+args.objectname+"']/../../"
        #print hostname
        for findrule in document.findall(hostname):#(".//rules/entry/source/[member%s" %hostname):
            rulename = "='"+findrule.attrib.get("name")+"']"
            if findrule.find(".//*[member='"+args.objectname+"']")!= None and findrule.find("destination")!=None:
                rulesfound.append(rulename)
            print rulename
        allobjects=False
            # if findrule.find(".//*[member='"+args.objectname+"']")!= None and findrule.getparent().tag=="address-group":
            #     group= findrule.attrib.get("name")
            #     rulegroup = "='"+findrule.attrib.get("name")+"']"
            #     print group
            #     groupname=".//*[member='"+findrule.attrib.get("name")+"']/../../"
            #     for findrule2 in document.findall(groupname):#(".//rules/entry/source/[member%s" %hostname):
            #         rulename2 = "='"+findrule2.attrib.get("name")+"']"
            #         if findrule2.find(".//*[member='"+findrule.attrib.get("name")+"']")!= None and findrule2.find("destination")!=None:
            #             print findrule2.getparent().tag
            #             test = findrule2.find(".//*[member='"+group+"']")
            #             test2 = findrule2.text
            #             print rulename2
            #             print test2
                        #rulesfound.append(rulename2)
                        #print rulename2
            
    print rulesfound
    if args.objectname == None and args.rulename!= None:
        rulename= "='"+args.rulename+"']"
        rulesfound.append(rulename)
        allobjects=False
    if args.objectname != None and args.rulename!= None:
        sys.exit("It is not possible to find by object and rule at the same time. Stopping execution")

    if args.excelname == None:
        workbook = xlsxwriter.Workbook('Firewall_Policies.xlsx') # Create Excel spreadsheet.
    else:
        workbook = xlsxwriter.Workbook(args.excelname) # Create Excel spreadsheet.

    worksheet = workbook.add_worksheet("Rules") # Create new worksheet within the spreadsheet.
    worksheet_objects = workbook.add_worksheet("Objects")

    bold = workbook.add_format({'bold': True}) # Cell formatting for row header

    dataformat = workbook.add_format() # Cell Formatting for data.
    dataformat.set_align('top')

    excelobj = Spreadsheet()
    excelobjects =Spreadsheet()
    excelobj.writeRowHeaders() # Create friendly row headers in the spreadsheet.
    excelobj.writeObjectHeaders()

    for numberrules in rulesfound:
        print numberrules
        for config in document: # Start after root (result)
            for devices in config.iter(virtual_or_deviceg): #depending of xml file value here can be "device-group" or "vsys1"
                for device in devices.findall(".//*[@name%s" %firewall_name):
                    for rulesection in device:#.iter("%s" %rule_type):
                        for rulestype in rulesection:#.iter("%s" %security_or_nat): # Start after result (rules)
                            for rules in rulestype.findall("rules"):
                                for entries in rules.findall(".//*[@name%s" %numberrules): # Start iterating after rules (entries)
                                    row += 1
                                    excelobj.setName(name=entries.attrib.get("name")) # Populate the rule description. Used attrib.get since name is a value within the tag.

                                    for fromzone in entries.findall("from"): # From zone block
                                        for members in fromzone.findall("member"): # From zone block - members block
                                            excelobj.setFromMember(members.text)

                                    for tozone in entries.findall("to"): # To zone block
                                        for members in tozone.findall("member"): # To zone block - members block
                                            excelobj.setToMember(members.text)

                                    for source in entries.findall("source"): # From source block
                                        for members in source.findall("member"): # From source block - members block
                                            excelobj.setSource(members.text)
                                            if allobjects==False:
                                                objectrow = findbyobjectname(document, members.text, objectrow)

                                    for source_trans in entries.findall("source-translation"): # From source block
                                        for members in source_trans: # From source block - members block
                                            if members.text is not None:
                                                excelobj.setSourceTranslation(members.text)
                                                # print "is not empty"
                                                # print members.text
                                            for submembers in members:
                                                if submembers.text is not None:
                                                    excelobj.setSourceTranslation(submembers.text)
                                                    # print "is not empty"
                                                    # print submembers.text
                                                for last in submembers:
                                                    if last.text is not None:
                                                        excelobj.setSourceTranslation(last.text)
                                                        # print "is not empty"
                                                        # print last.text

                                    for destination in entries.findall("destination"): # application block
                                        for members in destination.findall("member"): # application block - members block
                                            excelobj.setDestination(members.text)
                                            if allobjects== False:
                                                objectrow = findbyobjectname(document, members.text, objectrow)


                                    for destination_trans in entries.findall("destination-translation"): # From source block
                                        for members in destination_trans: # From source block - members block
                                            if members.text is not None:
                                                excelobj.setDestinationTranslation(members.text)
                                                # print "is not empty"
                                                # print members.text
                                            for submembers in members:
                                                if submembers.text is not None:
                                                    excelobj.setDestinationTranslation(submembers.text)
                                                    # print "is not empty"
                                                    # print submembers.text
                                                for last in submembers:
                                                    if last.text is not None:
                                                        excelobj.setDestinationTranslation(last.text)
                                                        # print "is not empty"
                                                        # print last.text

                                    for application in entries.findall("application"): # application block
                                        for members in application.findall("member"): # application block - members block
                                            excelobj.setApplication(members.text)

                                    for service in entries.findall("service"): # application block
                                        for members in service.findall("member"): # application block - members block
                                            excelobj.setServices(members.text)

                                    for hipprofiles in entries.findall("hip-profiles"): # application block
                                        for members in hipprofiles.findall("member"): # application block - members block
                                            excelobj.setHipprofiles(members.text)

                                    for action in entries.findall("action"):
                                        excelobj.setAction(action.text)

                                    for description in entries.findall("description"):
                                        excelobj.setDescription(description.text)

                                    for logstart in entries.findall("log-start"):
                                        excelobj.setLogstart(logstart.text)

                                    for logend in entries.findall("log-end"):
                                        excelobj.setLogend(logend.text)

                                    for tag in entries.findall("tag"): # application block
                                        for members in tag.findall("member"): # application block - members block
                                            excelobj.setTag(members.text)

                                    for profilesetting in entries.findall("profile-settings"): # application block
                                        for members in tag.findall("member"): # application block - members block
                                            excelobj.setTag(members.text)

                                    for disabled in entries.findall("disabled"):
                                        excelobj.setDisabled(disabled.text)

                                    for expiration in entries.findall("schedule"):
                                        excelobj.setExpiration(expiration.text)

                                    #rulesection
                                    excelobj.setRuletype(rulestype.tag)
                                    excelobj.setRulesection(rulesection.tag)
                                    excelobj.setFirewall(device.attrib.get("name")) 
                                    #print(device.attrib.get("name"))
                                    
                                    #firewall

                                    excelobj.writeRow(row) # Write each row to the spreadsheet.
                                    excelobj.newRow() # Clear old values and start new row.
    if allobjects==True:
        getObjects(document)
    workbook.close() # Close the spreadsheet since we are done with it now.


#TODO:
#Make it work with object groups
