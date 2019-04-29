try:
    with open("oid.txt") as file:
        oid = file.read()
        if oid:
            print("OID found: "+oid)
        else:
            oid = input("OID file empty, please enter next OID to be used: ")
except FileNotFoundError:
    oid = input("No OID file found, please enter next OID to be used: ")
    


oid = int(oid) + 1
with open("oid.txt", "w+") as file:
    file.write(str(oid))