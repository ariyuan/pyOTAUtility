from comtypes.client import CreateObject

# Login Credentials
qcServer = "http://<ALM server hostname>:8080/qcbin/"
qcUser = "sa"
qcPassword = ""
qcDomain = "Test"
projects = ["TestProj"]

# Connection
td = None


def login():
    global td
    for project in projects:
        # Do the actual login
        td = CreateObject("TDApiOle80.TDConnection")
        td.InitConnectionEx(qcServer)
        td.Login(qcUser, qcPassword)
        td.Connect(qcDomain, project)
        if td.Connected:
            print "System: Logged in to " + project
        else:
            print "Connect failed to " + project


def logout():
    global td
    if td.Connected:
        td.Disconnect
        td.Logout
        print "System: Logged out"
    td = None


def update_required_property(table_name, field_name, required):
    cust = td.Customization
    cust_fields = cust.Fields
    test_field = cust_fields.Field[table_name, field_name]
    if test_field.IsSupportsRequired:
        test_field.IsRequired = required
        cust.Commit()
        print "Require property for field {0} has been set to {1}".format(field_name, required)
    else:
        print "Require property for field {0} is not supported".format(field_name)


def update_field_type(table_name, field_name, field_type):
    cust = td.Customization
    cust_fields = cust.Fields
    test_field = cust_fields.Field[table_name, field_name]
    print "Current type of the field {0} is {1}".format(field_name, test_field.Type)
    test_field.Type = field_type
    cust.Commit()
    print "Current type of the field {0} has been changed to {1}".format(field_name, test_field.Type)


def update_assigned_list(table_name, field_name, new_list_name):
    cust = td.Customization
    cust_fields = cust.Fields
    test_field = cust_fields.Field[table_name, field_name]
    list = test_field.List
    print "Current list of the field {0} is {1}".format(field_name, test_field.List.Name)
    new_list = cust.Lists.List[new_list_name]
    test_field.List = new_list
    cust.Commit()
    print "Current list of the field {0} has been changed to {1}".format(field_name, test_field.List.Name)


def add_new_list(new_list_name, list_content):
    pass


def list_item_update(list_name, old_item_name, new_item_name):
    cust = td.Customization
    list = cust.Lists.List[list_name]
    root = list.RootNode
    child = root.Child[old_item_name]
    print "Original node name is {0}".format(child.Name)
    child.Name = new_item_name
    cust.Commit()
    print "Node name has been changed to {0}".format(child.Name)


def update_field_length(field_name, length):
    pass

if __name__ == "__main__":
    login()
    list_item_update("TestRoot", "ItemChanged", "Item1")
    logout()

