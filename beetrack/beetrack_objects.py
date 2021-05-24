import sys
import json


class Item:
    def __init__(
        self, itemDict, name=None, description=None, quantity=None, code=None,
    ):
        if type(itemDict) == dict:
            self.id = itemDict["id"]
            self.name = itemDict["name"]
            self.description = itemDict["description"]
            self.quantity = itemDict["quantity"]
            self.code = itemDict["code"]
        elif type(itemDict) == str:
            self.id = itemDict
            self.name = name
            self.description = description
            self.quantity = quantity
            self.code = code
        else:
            raise TypeError(
                "Items must be created from dict or str, not {}".format(
                    type(type(itemDict))
                )
            )
        pass

    def dump_dict(self):
        itemDict = {}
        if self.id:
            itemDict["id"] = self.id
        if self.name:
            itemDict["name"] = self.name
        if self.description:
            itemDict["description"] = self.description
        if self.quantity:
            itemDict["quantity"] = self.quantity
        if self.code:
            itemDict["code"] = self.code
        return itemDict

    def dump_json(self):
        json.dumps(self.dump_dict())


dispatchTypeDict = {"LAST MILE": 0, "FIRST MILE": 1, "FULFILLMENT": 2}
dispatchTypeReverseDict = {v: k for k, v in dispatchTypeDict.items()}
dispatchPriorityDict = {"NORMAL": 0, "URGENTE": 1}
dispatchPriorityReverseDict = {v: k for k, v in dispatchPriorityDict.items()}


class Dispatch:
    def __init__(
        self,
        dispatchDict,
        contactName=None,
        contactAddress=None,
        contactPhone=None,
        contactEmail=None,
        contactID=None,
        contactComment=None,
        pickupAddress=None,
        priority=0,  # 0 = Normal, 1 = Urgente
        maxDeliveryTime=None,
        items=[],
        client=None,
        mode=3,
        dispatchType=0,  # 0 = Last Mile, 1 = First Mile, 2 = Fullfillment.
    ):
        if type(dispatchDict) == dict:
            self.id = dispatchDict["identifier"]
            self.contactName = dispatchDict["contact_name"]
            self.contactAddress = dispatchDict["contact_address"]
            self.contactPhone = dispatchDict["contact_phone"]
            self.contactID = dispatchDict["contact_id"]
            self.contactEmail = dispatchDict["contact_email"]
            # Iterar sobre tags:
            self.client = None
            self.contactComment = None
            self.priority = 0
            self.dispatchType = 0
            for tag in dispatchDict["tags"]:
                if tag["name"] == "Cliente":
                    self.client = tag["value"]
                elif tag["name"] == "Información adicional":
                    self.contactComment = tag["value"]
                elif tag["name"] == "Prioridad":
                    try:
                        self.priority = dispatchPriorityDict[tag["value"].upper()]
                    except KeyError:
                        print(
                            'WARNING: Unknown priority "{}". Setting normal priority.'.format(
                                tag["value"]
                            ),
                            file=sys.stderr,
                        )
                        self.priority = 0
                elif tag["name"] == "Tipo de despacho":
                    try:
                        self.dispatchType = dispatchTypeDict[tag["value"].upper()]
                    except KeyError:
                        print(
                            'ERROR: Unknown dispatch type "{}". Setting Last Mile.'.format(
                                tag["value"]
                            ),
                            file=sys.stderr,
                        )
                else:
                    print(
                        'ERROR: Unknown tag "{}". Discarding.'.format(tag["name"]),
                        file=sys.stderr,
                    )
            # Iterar sobre items
            self.items = []
            for itemDict in dispatchDict["items"]:
                self.items.append(Item(itemDict))
            self.pickupAddress = None  # Este no aparece en el API
        elif type(dispatchType) == str:
            self.id = dispatchDict
            self.contactName = contactName
            self.contactAddress = contactAddress
            self.contactPhone = contactPhone
            self.contactEmail = contactEmail
            self.contactID = contactID
            self.contactComment = contactComment
            self.pickupAddress = pickupAddress
            if type(priority) == int:
                self.priority = priority
            elif type(priority) == str:
                try:
                    self.priority = dispatchPriorityDict[priority.upper()]
                except KeyError:
                    print(
                        f'ERROR: Unknown priority "{priority}". Setting normal priority',
                        file=sys.stderr,
                    )
                    self.priority = 0
            else:
                print(
                    "ERROR: Priority must be either int or str, not {}. Setting normal priority.".format(
                        type(priority)
                    ),
                    file=sys.stderr,
                )
                self.priority = 0
            self.maxDeliveryTime = maxDeliveryTime
            self.items = items
            self.client = client
            self.mode = mode
            if type(dispatchType) == int:
                self.dispatchType = dispatchType
            elif type(dispatchType) == str:
                try:
                    self.dispatchType = dispatchPriorityDict[dispatchType.upper()]
                except KeyError:
                    print(
                        f'ERROR: Unknown dispatch type "{dispatchType}". Setting Last Mile.',
                        file=sys.stderr,
                    )
                    self.dispatchType = 0
            else:
                print(
                    "ERROR: Dispatch type must be either int or str, not {}. Setting Last Mile.".format(
                        type(dispatchType)
                    )
                )
                self.dispatchType = 0
        else:
            raise TypeError(
                "Dispatches must be created with a dict or a str, not {}".format(
                    type(dispatchDict)
                )
            )
        if id == None:
            print("ERROR: Dispatches must have an ID")
            raise Exception("Dispatches must have an ID")
        pass

    def dump_dict(self):
        priorityStr = dispatchPriorityReverseDict[self.priority]
        dispatchTypeStr = dispatchTypeReverseDict[self.dispatchType]
        dispatchDict = {
            "identifier": self.id,
            "tags": [
                {"name": "Prioridad", "value": priorityStr.title()},
                {"name": "Tipo de despacho", "value": dispatchTypeStr.title()},
            ],
        }
        if self.contactName:
            dispatchDict["contact_name"] = self.contactName
        if self.contactAddress:
            dispatchDict["contact_address"] = self.contactAddress
        if self.contactPhone:
            dispatchDict["contact_phone"] = self.contactPhone
        if self.contactID:
            dispatchDict["contact_id"] = self.contactID
        if self.contactEmail:
            dispatchDict["contact_email"] = self.contactEmail
        if self.contactComment:
            dispatchDict["tags"].append(
                {"name": "Información adicional", "value": self.contactComment}
            )
        if len(self.items) >= 1:
            dispatchDict["items"] = []
            for item in self.items:
                dispatchDict["items"].append(item.dump_dict())
        if self.client:
            dispatchDict["tags"].append({"name": "Cliente", "value": self.client})
        if self.pickupAddress:
            dispatchDict["pickup_address"] = {"name": self.pickupAddress}
        return dispatchDict

    def dump_json(self):
        return json.dumps(self.dump_dict())
