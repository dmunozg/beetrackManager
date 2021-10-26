import sys
import json
from six import string_types


class Item:
    weight = None

    def __init__(
        self,
        itemDict=None,
        name=None,
        description=None,
        quantity=0,
        code=None,
    ):
        if isinstance(itemDict, dict):
            self.id = itemDict["id"]
            self.name = itemDict["name"]
            self.description = itemDict["description"]
            self.quantity = itemDict["quantity"]
            self.code = itemDict["code"]
        elif isinstance(code, str):
            self.id = itemDict
            self.name = name
            self.description = description
            self.quantity = quantity
            self.code = code
        else:
            raise TypeError(
                "Items must be created from dict or str, not {}".format(type(itemDict))
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
        if self.quantity is not None:
            itemDict["quantity"] = self.quantity
        if self.code:
            itemDict["code"] = self.code
        if self.weight:
            itemDict["extras"] = [{"name": "Peso", "value": str(self.weight)}]
        return itemDict

    def dump_json(self):
        json.dumps(self.dump_dict())


# TODO
# Esto hay que cambiarlo por una base de datos
dispatchTypeDict = {"LAST MILE": 0, "FIRST MILE": 1, "FULFILLMENT": 2, "FORWARDING": 3}
dispatchTypeReverseDict = {v: k for k, v in dispatchTypeDict.items()}
dispatchPriorityDict = {"NORMAL": 0, "URGENTE": 1}
dispatchPriorityReverseDict = {v: k for k, v in dispatchPriorityDict.items()}


class Dispatch:
    distributionCenter = None
    bulkType = None
    admission = None
    forwardingSender = None
    forwardingSenderAddress = None

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
        firstMileDestination=None,
        additionalDocument=None,
        priority=0,  # 0 = Normal, 1 = Urgente
        maxDeliveryTime=None,
        minDeliveryTime=None,
        items=[],
        client=None,
        mode=2,
        dispatchType=0,  # 0 = Last Mile, 1 = First Mile, 2 = Fullfillment 3 = Forwarding.
    ):
        if isinstance(dispatchDict, dict):
            self.id = dispatchDict["identifier"]
            self.contactName = dispatchDict["contact_name"]
            self.contactAddress = dispatchDict["contact_address"]
            self.contactPhone = dispatchDict["contact_phone"]
            self.contactID = dispatchDict["contact_id"]
            self.contactEmail = dispatchDict["contact_email"]
            self.mindeliveryTime = dispatchDict["min_delivery_time"]
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
                elif tag["name"] == "FM_Direccion":
                    self.firstMileDestination = tag["value"]
                elif tag["name"] == "Documento adicional":
                    self.additionalDocument = tag["value"]
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
        elif isinstance(dispatchDict, string_types) or isinstance(dispatchDict, int):
            self.id = dispatchDict
            self.contactName = contactName
            self.contactAddress = contactAddress
            self.contactPhone = str(contactPhone)
            self.contactEmail = contactEmail
            self.contactID = contactID
            self.contactComment = contactComment
            self.pickupAddress = pickupAddress
            self.additionalDocument = additionalDocument
            self.firstMileDestination = firstMileDestination
            if type(priority) == int:
                self.priority = priority
            elif type(priority) == string_types:
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
            self.mindeliveryTime = minDeliveryTime
            self.items = items
            self.client = client
            self.mode = mode
            if type(dispatchType) == int:
                self.dispatchType = dispatchType
            elif type(dispatchType) == string_types:
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
        dispatchDict = {"identifier": self.id}
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
        if len(self.items) >= 1:
            dispatchDict["items"] = []
            for item in self.items:
                dispatchDict["items"].append(item.dump_dict())
        dispatchDict["tags"] = [
            {"name": "Prioridad", "value": priorityStr.title()},
            {"name": "Tipo de despacho", "value": dispatchTypeStr.title()},
        ]
        if self.contactComment:
            dispatchDict["tags"].append(
                {"name": "Información adicional", "value": self.contactComment}
            )
        if self.firstMileDestination:
            dispatchDict["tags"].append(
                {"name": "FM_Direccion", "value": self.firstMileDestination}
            )
        if self.client:
            dispatchDict["tags"].append({"name": "Cliente", "value": self.client})
        if self.additionalDocument:
            dispatchDict["tags"].append(
                {"name": "Documento adicional", "value": self.additionalDocument}
            )
        if self.pickupAddress:
            dispatchDict["pickup_address"] = {"name": self.pickupAddress}
        # Si el despacho es un Forwarding, debe incluir el centro de despacho.
        if self.dispatchType == 3:
            self.mode = 2
            dispatchDict["place"] = self.distributionCenter
            # Además debe incluír si fué un retiro o se despachó en bodega
            dispatchDict["tags"].append({"name": "Admisión", "value": self.admission})
            # Y marcar que ya fue recibido en bodega
            dispatchDict["status_id"] = 2
            # dispatchDict["substatus"] = "En bodega"
        if self.forwardingSender:
            dispatchDict["tags"].append(
                {"name": "FW_Remitente", "value": self.forwardingSender}
            )
        if self.forwardingSenderAddress:
            dispatchDict["tags"].append(
                {
                    "name": "FW_Remitente_Dirección",
                    "value": self.forwardingSenderAddress,
                }
            )
        if self.bulkType:
            dispatchDict["tags"].append(
                {"name": "Tipo de carga", "value": self.bulkType}
            )
        dispatchDict["mode"] = self.mode
        return dispatchDict

    def dump_json(self):
        return json.dumps(self.dump_dict(), ensure_ascii=False)
