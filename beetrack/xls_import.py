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
    "Sin documento": "ND",
}
dispatchTypeDict = {"LAST MILE": 0, "FIRST MILE": 1, "FULFILLMENT": 2}
dispatchPriorityDict = {"NORMAL": 0, "URGENTE": 1}


def excelrow_to_dispatch(excelRow, client, pickupAddress):
    """ Toma una fila de celdas excel con datos de despacho (en forma de tupla de celdas) y genera un objeto Dispatch a partir de este.
    La función devuelve una tupla con:
    (Dispatch resultante, código de error, lista de observaciones)
    
    el código de error será:
    0 : Sin errores
    1 : Errores no críticos
    2 : Errores críticos. Se devuelve un Dispatch=None en este caso
    
    la lista de observaciones será una lista de str con la información erronea o faltante del despacho ingresado."""
    warningList = []  # Lista de observaciones
    errorCode = 0
    # Validadores
    def _validate_ID(cell):
        nonlocal errorCode
        if cell.value == None:
            errorCode = 2
            warningList.append("Crítico: No puede haber un despacho sin código.")
        else:
            return cell.value

    def _validate_document_type(cell):
        nonlocal errorCode
        if cell.value in documentTypeDict.keys():
            _documentType = documentTypeDict[cell.value]
        elif cell.value == None:
            errorCode = 1 if errorCode != 2 else 2
        else:
            errorCode = 1 if errorCode != 2 else 2
            warningList.append(
                'Error: No se reconoce el tipo de documento "{}". Se dejará como "Otro" si no se corrige.'.format(
                    cell.value
                )
            )
            _documentType = "O"
        return _documentType

    def _validate_document_number(cell, docType):
        nonlocal errorCode
        if docType != "ND":
            _documentNumber = f"{docType} {cell.value}"
        else:
            _documentNumber = docType
        return _documentNumber

    def _validate_item_quantity(cell):
        nonlocal errorCode
        if cell.value == None:
            errorCode = 1 if errorCode != 2 else 2
            warningList.append(
                "Advertencia: No se especificó un número de bultos. Se dejará en 0 si no se corrige"
            )
            _itemQuantity = 0
        elif type(cell.value) != int:
            try:
                _itemQuantity = int(cell.value)
                errorCode = 1 if errorCode != 2 else 2
                warningList.append(
                    "Advertencia: El número de bultos debe ser un valor entero. Se registraron {} bulto(s)".format(
                        _itemQuantity
                    )
                )
            except ValueError:
                errorCode = 2
                warningList.append(
                    "Crítico: El número de bultos debe ser un valor entero."
                )
                _itemQuantity = None
        elif cell.value < 0:
            errorCode = 2
            warningList.append("Crítico: El número de bultos no puede ser negativo.")
            _itemQuantity = None
        else:
            _itemQuantity = cell.value
        return _itemQuantity

    def _validate_transport_type(cell):
        nonlocal errorCode
        if cell.value == None:
            errorCode = 1 if errorCode != 2 else 2
            warningList.append(
                "Advertencia: No se especificó un tipo de transporte. Se dejará como Last Mile si no se corrige."
            )
            _transportType = 0
        elif type(cell.value) != str:
            errorCode = 1 if errorCode != 2 else 2
            warningList.append(
                "Advertencia: No se pudo interpretar el tipo de transporte. Verifique si la casilla está en formato de texto. Se dejará como Last Mile si no se corrige."
            )
            _transportType = 0
        elif cell.value.upper() not in dispatchTypeDict.keys():
            errorCode = 1 if errorCode != 2 else 2
            warningList.append(
                'Advertencia: "{}" no se reconoce como un tipo de transporte. Se dejará como Last Mile si no se corrige'.format(
                    cell.value
                )
            )
            _transportType = 0
        else:
            _transportType = dispatchTypeDict[cell.value.upper()]
        return _transportType

    def _validate_contact_address(cell1, cell2):
        nonlocal errorCode
        return f"{cell1.value}, {cell2.value}"

    def _validate_priority(cell):
        nonlocal errorCode
        if cell.value == None:
            _priority = 0
        elif type(cell.value) != str:
            errorCode = 1 if errorCode != 2 else 2
            warningList.append(
                "Advertencia: No se reconoce la prioridad. Se dejará normal si no se corrige."
            )
            _priority = 0
        elif cell.value.upper() not in dispatchPriorityDict.keys():
            errorCode = 1 if errorCode != 2 else 2
            warningList.append(
                f'Advertencia: No se reconoce "{cell.value}" como prioridad. Se dejará normal si no se corrige.'
            )
            _priority = 0
        else:
            _priority = dispatchPriorityDict[cell.value.upper()]
        return _priority

    # Importar cada celda
    dispatchID = _validate_ID(excelRow[0])
    documentType = _validate_document_type(excelRow[1])
    documentNumber = _validate_document_number(excelRow[2], documentType)
    additionalDocument = excelRow[3].value
    itemDescription = excelRow[4].value
    itemQuantity = _validate_item_quantity(excelRow[5])
    transportType = _validate_transport_type(excelRow[6])
    contactName = excelRow[7].value
    contactPhone = excelRow[8].value
    contactEmail = excelRow[9].value
    contactAddress = _validate_contact_address(excelRow[10], excelRow[11])
    contactComment = excelRow[12].value
    maxDeliveryTime = excelRow[13].value
    priority = _validate_priority(excelRow[14])
    firstMileTransporter = excelRow[15].value
    contactID = excelRow[16].value
    # Si el despacho es First Mile la dirección debe ser la dirección del transporte, y el destinatario debe ir en firstMileDestination
    if transportType == 1:
        firstMileDestination = contactAddress
        contactAddress = firstMileTransporter
    else:
        firstMileDestination = None
    # Generar el output sólo si no hubieron errores críticos
    if errorCode == 2:
        resultingDispatch = None
    else:
        # Generar el objeto Item que describe los bultos
        resultingItem = Item(
            description=itemDescription, quantity=itemQuantity, code=documentNumber
        )
        # Generara el objeto Dispatch
        resultingDispatch = Dispatch(
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
            client=client,
            firstMileDestination=firstMileDestination,
            pickupAddress=pickupAddress,
            items=[resultingItem],
            additionalDocument=additionalDocument,
        )
    return (resultingDispatch, errorCode, warningList)


def xlsx_to_dispatches(xlsxFilename, client, pickupAddress):
    warningSet = set()
    foundDispatches = []
    try:
        xlsxData = openpyxl.load_workbook(xlsxFilename, data_only=True, read_only=True)
    except:
        warningSet.add("No fue posible abrir el archivo .xlsx. Posiblemente corrupto.")
        return (foundDispatches, warningSet)
    dispatchesSheet = xlsxData.active
    for row in dispatchesSheet.iter_rows(min_row=2, max_col=16):
        # Omitir la fila si está vacía
        if not any(cell.value for cell in row):
            continue
        # Revisar si la fila tiene un código de despacho, omitir la fila si no lo tiene.
        elif row[0].value == None:
            warningSet.add(
                "Se encontraron filas sin código de despacho que fueron omitidas."
            )
            continue
        else:
            foundDispatches.append(excelrow_to_dispatch(row, client, pickupAddress))
    return (foundDispatches, warningSet)
