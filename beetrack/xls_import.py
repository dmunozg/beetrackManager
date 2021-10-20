from os import error
from .beetrack_objects import Dispatch, Item
import openpyxl

# TODO
# Esto hay que cambiarlo por una base de datos
documentTypeDict = {
    "Factura": "F",
    "Boleta": "B",
    "Guía de despacho": "G",
    "Otro": "O",
    "Orden de compra": "OC",
    "Sin documento": "ND",
}
dispatchTypeDict = {"LAST MILE": 0, "FIRST MILE": 1, "FULFILLMENT": 2, "FORWARDING": 3}
dispatchPriorityDict = {"NORMAL": 0, "URGENTE": 1}


class defaultRowParser:
    def __init__(self, excelRow, client, pickupAddress):
        self.excelRow = excelRow
        self.client = client
        self.pickupAddress = pickupAddress
        self.resultingDispatch = None
        self.warningList = []
        self.errorCode = 0

    # Validadores
    def validate_ID(self, cell):
        if cell.value == None:
            self.errorCode = 2
            self.warningList.append("Crítico: No puede haber un despacho sin código.")
            return "NONE"
        elif not isinstance(cell.value, str):
            self.errorCode = 2
            self.warningList.append("Crítico: El código debe ser alfanumérico.")
            return "FAIL"
        else:
            return cell.value

    def validate_document_type(self, cell):
        if cell.value == None:
            self.errorCode = 1 if self.errorCode != 2 else 2
            documentType = "O"
            self.warningList.append(
                'No se especificó un tipo de documento. Se dejará como "Otro" si no se corrige.'.format(
                    cell.value
                )
            )
        else:
            for dt in documentTypeDict.keys():
                if cell.value.upper().strip() == dt.upper():
                    documentType = documentTypeDict[dt]
                    break
            else:
                self.errorCode = 1 if self.errorCode != 2 else 2
                self.warningList.append(
                    'No se reconoce el tipo de documento "{}". Se dejará como "Otro" si no se corrige.'.format(
                        cell.value
                    )
                )
                documentType = "O"
        return documentType

    def validate_document_number(self, cell, docType):
        if docType != "ND":
            documentNumber = f"{docType} {cell.value}"
        else:
            documentNumber = docType
        return documentNumber

    def validate_item_quantity(self, cell):
        if cell.value == None:
            self.errorCode = 1 if self.errorCode != 2 else 2
            self.warningList.append(
                "No se especificó un número de bultos. Se dejará en 0 si no se corrige"
            )
            itemQuantity = 0
        elif type(cell.value) != int:
            try:
                itemQuantity = int(cell.value)
                if cell.value % 1 != 0:
                    self.errorCode = 1 if self.errorCode != 2 else 2
                    self.warningList.append(
                        "El número de bultos debe ser un valor entero. Se registraron {} bulto(s)".format(
                            itemQuantity
                        )
                    )
            except ValueError:
                self.errorCode = 2
                self.warningList.append("El número de bultos debe ser un valor entero.")
                itemQuantity = None
        elif cell.value < 0:
            self.errorCode = 2
            self.warningList.append("El número de bultos no puede ser negativo.")
            itemQuantity = None
        else:
            itemQuantity = cell.value
        return itemQuantity

    def validate_transport_type(self, cell):
        if cell.value == None:
            self.errorCode = 1 if self.errorCode != 2 else 2
            self.warningList.append(
                "No se especificó un tipo de transporte. Se dejará como Last Mile si no se corrige."
            )
            transportType = 0
        elif type(cell.value) != str:
            self.errorCode = 1 if self.errorCode != 2 else 2
            self.warningList.append(
                "No se pudo interpretar el tipo de transporte. Verifique si la casilla está en formato de texto. Se dejará como Last Mile si no se corrige."
            )
            transportType = 0
        elif cell.value.upper() not in dispatchTypeDict.keys():
            self.errorCode = 1 if self.errorCode != 2 else 2
            self.warningList.append(
                '"{}" no se reconoce como un tipo de transporte. Se dejará como Last Mile si no se corrige'.format(
                    cell.value
                )
            )
            transportType = 0
        else:
            transportType = dispatchTypeDict[cell.value.upper()]
        return transportType

    def validate_contact_address(self, cell1, cell2):
        return f"{cell1.value}, {cell2.value}"

    def validate_priority(self, cell):
        if cell.value == None:
            priority = 0
        elif type(cell.value) != str:
            self.errorCode = 1 if self.errorCode != 2 else 2
            self.warningList.append(
                "No se reconoce la prioridad indicada. Se dejará normal si no se corrige."
            )
            priority = 0
        elif cell.value.upper() not in dispatchPriorityDict.keys():
            self.errorCode = 1 if self.errorCode != 2 else 2
            self.warningList.append(
                f'No se reconoce "{cell.value}" como prioridad. Se dejará normal si no se corrige.'
            )
            priority = 0
        else:
            priority = dispatchPriorityDict[cell.value.upper()]
        return priority

    def parse(self):
        dispatchID = self.validate_ID(self.excelRow[0])
        documentType = self.validate_document_type(self.excelRow[1])
        documentNumber = self.validate_document_number(self.excelRow[2], documentType)
        additionalDocument = self.excelRow[3].value
        itemDescription = self.excelRow[4].value
        itemQuantity = self.validate_item_quantity(self.excelRow[5])
        transportType = self.validate_transport_type(self.excelRow[6])
        contactName = self.excelRow[7].value
        contactPhone = self.excelRow[8].value
        contactEmail = self.excelRow[9].value
        contactAddress = self.validate_contact_address(
            self.excelRow[10], self.excelRow[11]
        )
        contactComment = self.excelRow[12].value
        maxDeliveryTime = self.excelRow[13].value
        priority = self.validate_priority(self.excelRow[14])
        firstMileTransporter = self.excelRow[15].value
        contactID = self.excelRow[16].value
        # Si el despacho es First Mile la dirección debe ser la dirección del transporte, y el destinatario debe ir en firstMileDestination
        if transportType == 1:
            firstMileDestination = contactAddress
            contactAddress = firstMileTransporter
        else:
            firstMileDestination = None
        # Generar el output sólo si no hubieron errores críticos
        if self.errorCode == 2:
            self.resultingDispatch = dispatchID
            return 1
        else:
            # Generar el objeto Item que describe los bultos
            resultingItem = Item(
                description=itemDescription, quantity=itemQuantity, code=documentNumber
            )
            # Generara el objeto Dispatch
            self.resultingDispatch = Dispatch(
                dispatchID,
                contactName=contactName,
                contactAddress=contactAddress,
                contactPhone=contactPhone,
                contactEmail=contactEmail,
                contactID=contactID,
                contactComment=contactComment,
                priority=priority,
                maxDeliveryTime=maxDeliveryTime,
                dispatchType=transportType,
                client=self.client,
                firstMileDestination=firstMileDestination,
                pickupAddress=self.pickupAddress,
                items=[resultingItem],
                additionalDocument=additionalDocument,
            )
            return 0


class XlsxParser:
    warningSet = set()
    foundDispatches = []
    rowParser = defaultRowParser
    firstRowToCheck = 2
    lastColumnToCheck = 17
    dispatchIDcolumnIndex = 0

    def __init__(self, xlsxFilename, client, pickupAddress):
        self.xlsxFilename = xlsxFilename
        self.client = client
        self.pickupAddress = pickupAddress
        self.warningSet = set()
        self.foundDispatches = []
        self.rowParser = self.rowParser
        self.firstRowToCheck = self.firstRowToCheck
        self.lastColumnToCheck = self.lastColumnToCheck
        self.dispatchIDcolumnIndex = self.dispatchIDcolumnIndex

    def parse(self):
        try:
            xlsxData = openpyxl.load_workbook(
                self.xlsxFilename, data_only=True, read_only=True
            )
        except:
            self.warningSet.add(
                "No fue posible abrir el archivo .xlsx, posiblemente corrupto. Verifique que puede abrirlo con Excel y envíelo denuevo."
            )
            return 1
        dispatchesSheet = xlsxData.active
        for row in dispatchesSheet.iter_rows(
            min_row=self.firstRowToCheck, max_col=self.lastColumnToCheck
        ):
            # Omitir la fila si está vacía
            if not any(cell.value for cell in row):
                continue
            # Revisar si la fila tiene un código de despacho, omitir la fila si no lo tiene.
            elif row[self.dispatchIDcolumnIndex].value == None:
                self.warningSet.add(
                    "Se encontraron filas sin código de despacho que fueron omitidas. Todos los despachos deben incluir uno."
                )
                continue
            elif not isinstance(row[self.dispatchIDcolumnIndex].value, str):
                self.warningSet.add("Se encontraron filas con códigos no-alfanuméricos")
                continue
            else:
                scannedRow = self.rowParser(row, self.client, self.pickupAddress)
                scannedRow.parse()
                self.foundDispatches.append(
                    (
                        scannedRow.resultingDispatch,
                        scannedRow.errorCode,
                        scannedRow.warningList,
                    )
                )


class BbvinosRowParser(defaultRowParser):
    def parse(self):
        dispatchID = self.validate_ID(self.excelRow[1])
        documentType = self.validate_document_type(self.excelRow[2])
        documentNumber = self.validate_document_number(self.excelRow[3], documentType)
        additionalDocument = self.excelRow[4].value
        itemDescription = self.excelRow[5].value
        itemQuantity = self.validate_item_quantity(self.excelRow[6])
        transportType = self.validate_transport_type(self.excelRow[7])
        contactName = self.excelRow[8].value
        contactPhone = self.excelRow[9].value
        contactEmail = self.excelRow[10].value
        contactAddress = self.validate_contact_address(
            self.excelRow[11], self.excelRow[12]
        )
        contactComment = self.excelRow[13].value
        maxDeliveryTime = self.excelRow[14].value
        priority = self.validate_priority(self.excelRow[15])
        firstMileTransporter = self.excelRow[16].value
        contactID = self.excelRow[17].value
        # Si el despacho es First Mile la dirección debe ser la dirección del transporte, y el destinatario debe ir en firstMileDestination
        if transportType == 1:
            firstMileDestination = contactAddress
            contactAddress = firstMileTransporter
        else:
            firstMileDestination = None
        # Generar el output sólo si no hubieron errores críticos
        if self.errorCode == 2:
            self.resultingDispatch = dispatchID
            return 1
        else:
            # Generar el objeto Item que describe los bultos
            resultingItem = Item(
                description=itemDescription, quantity=itemQuantity, code=documentNumber
            )
            # Generara el objeto Dispatch
            self.resultingDispatch = Dispatch(
                dispatchID,
                contactName=contactName,
                contactAddress=contactAddress,
                contactPhone=contactPhone,
                contactEmail=contactEmail,
                contactID=contactID,
                contactComment=contactComment,
                priority=priority,
                maxDeliveryTime=maxDeliveryTime,
                dispatchType=transportType,
                client=self.client,
                firstMileDestination=firstMileDestination,
                pickupAddress=self.pickupAddress,
                items=[resultingItem],
                additionalDocument=additionalDocument,
            )
            return 0


class BbvinosXlsxParser(XlsxParser):
    rowParser = BbvinosRowParser
    firstRowToCheck = 2
    lastColumnToCheck = 18
    dispatchIDcolumnIndex = 1


admissionTypeSet = set(["Retiro", "CD Renca", "CD Puente Alto"])
bulkTypeSet = set(["Seca", "Peligrosa", "Refrigerada"])


class BodegaRowParser(defaultRowParser):
    def validate_admission(self, cell):
        if cell.value == None:
            self.errorCode = 1 if self.errorCode != 2 else 2
            self.warningList.append(
                "No se especificó el tipo de admisión. Se dejará como Retiro"
            )
            admissionType = "Retiro"
        elif type(cell.value) != str:
            self.errorCode = 1 if self.errorCode != 2 else 2
            self.warningList.append(
                "No se pudo interpretar el tipo de admisión. Se dejará como Retiro"
            )
            admissionType = "Retiro"
        elif cell.value.upper() not in [i.upper() for i in admissionTypeSet]:
            self.warningList.append(
                "{} no está en la lista de centros de distribución. Se dejará como retiro si no se corrije.".format(
                    cell.value
                )
            )
            admissionType = "Retiro"
        else:
            for cd in admissionTypeSet:
                if cell.value.upper() == cd.upper():
                    admissionType = cd
        return admissionType

    def validate_bulk_type(self, cell):
        if cell.value == None:
            self.errorCode = 1 if self.errorCode != 2 else 2
            self.warningList.append(
                "No se especificó el tipo de carga. Se dejará como carga seca"
            )
            bulkType = "Seca"
        elif type(cell.value) != str:
            self.errorCode = 1 if self.errorCode != 2 else 2
            self.warningList.append(
                "No se pudo interpretar el tipo de carga. Se dejará como carga seca"
            )
            bulkType = "Seca"
        else:
            for bt in bulkTypeSet:
                if bt.upper() in cell.value.upper().strip():
                    bulkType = bt
                    break
            else:
                self.errorCode = 1 if self.errorCode != 2 else 2
                self.warningList.append(
                    '"{}" no se reconoce como un tipo de carga. Se dejará como carga seca si no se corrije.'.format(
                        cell.value
                    )
                )
                bulkType = "Normal"
        return bulkType

    def validate_weight(self, cell):
        if cell.value == None:
            self.errorCode = 2
            self.warningList.append(
                "Error! Incluya el peso en kilogramos de los bultos."
            )
            return None
        else:
            try:
                weight = float(cell.value)
            except ValueError:
                self.errorCode = 2
                self.warningList.append(
                    "Error! El peso de los bultos debe ser un número."
                )
                return None
        return weight

    def parse(self):
        dispatchID = self.validate_ID(self.excelRow[0])
        documentType = self.validate_document_type(self.excelRow[1])
        documentNumber = self.validate_document_number(self.excelRow[2], documentType)
        additionalDocument = self.excelRow[3].value
        itemDescription = self.excelRow[4].value
        itemQuantity = self.validate_item_quantity(self.excelRow[17])
        transportType = self.validate_transport_type(self.excelRow[5])
        contactName = self.excelRow[6].value
        contactPhone = self.excelRow[7].value
        contactEmail = self.excelRow[8].value
        contactAddress = self.validate_contact_address(
            self.excelRow[10], self.excelRow[9]
        )
        contactComment = self.excelRow[11].value
        priority = self.validate_priority(self.excelRow[12])
        # Validar datos de bodega
        senderName = self.excelRow[14].value
        try:
            senderAddress = self.validate_contact_address(
                self.excelRow[16], self.excelRow[15]
            )
        except:
            senderAddress = None
        admission = self.validate_admission(self.excelRow[19])
        weight = self.validate_weight(self.excelRow[18])
        bulkType = self.validate_bulk_type(self.excelRow[20])
        # Generar el output sólo si no hubieron errores críticos
        if self.errorCode == 2:
            self.resultingDispatch = dispatchID
            return 1
        else:
            # Generar el objeto Item que describe los bultos
            resultingItem = Item(
                description=itemDescription, quantity=itemQuantity, code=documentNumber
            )
            resultingItem.weight = weight
            # Generara el objeto Dispatch
            self.resultingDispatch = Dispatch(
                dispatchID,
                contactName=contactName,
                contactAddress=contactAddress,
                contactPhone=contactPhone,
                contactEmail=contactEmail,
                contactComment=contactComment,
                priority=priority,
                dispatchType=transportType,
                client=self.client,
                items=[resultingItem],
                additionalDocument=additionalDocument,
            )
            self.resultingDispatch.admission = admission
            if admission != "Retiro":
                self.resultingDispatch.distributionCenter = admission
            else:
                self.resultingDispatch.distributionCenter = "CD Renca"
            self.resultingDispatch.bulkType = bulkType
            self.resultingDispatch.forwardingSender = senderName
            self.resultingDispatch.forwardingSenderAddress = senderAddress
            return 0


class BodegaXlsxParser(XlsxParser):
    rowParser = BodegaRowParser
    lastColumnToCheck = 21
