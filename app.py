from beetrack.beetrack_api import BeetrackAPI
from beetrack import xls_import, mail_handler
import os, time, sys, json
import json, os, time, sys
import pandas as pd

REQUIRED_ENV_VARIABLES = [
    "BEETRACK_APIKEY",
    "IMAP_USER",
    "IMAP_PASSWD",
    "IMAP_SERVER",
    "SMTP_SERVER",
    "SMTP_PORT",
]


def check_if_env_is_set(env):
    if os.getenv(env):
        return True
    else:
        return False


if not all([check_if_env_is_set(env) for env in REQUIRED_ENV_VARIABLES]):
    print(
        "[{timestamp}] Not all required environment variables were set before runtime. Loading from .env file.".format(
            timestamp=time.strftime("%H:%M:%S")
        )
    )
    from dotenv import load_dotenv

    load_dotenv()


def check_env_definitive(env):
    if not os.getenv(env):
        print(
            "[{timestamp}] Critical: Environment variable {env} was not set.".format(
                timestamp=time.strftime("%H:%M:%S"), env=env
            ),
            file=sys.stderr,
        )
        sys.exit(1)


for env in REQUIRED_ENV_VARIABLES:
    check_env_definitive(env)

BASE_URL = "https://app.beetrack.com/api/external/v1"
ALLOWED_FILES_EXTENSIONS = [
    "XLSX",
]
# TODO
# Esto hay que reemplazarlo por una tabla SQL
ALLOWED_CLIENTS_DF = pd.DataFrame(
    [
        [
            "lkeeler.flipout@gmail.com",
            "Pruebas",
            "TEST",
            "Reibo 3619, Puente Alto",
            True,  # Allow override
            None,  # Custom Parser
        ],
        [
            "matias@logicaexpress.cl",
            "Pruebas",
            "TEST",
            "Cerro Loma Larga 3624, Puente Alto",
            True,  # Allow override
            None,  # Custom Parser
        ],
        [
            "carolina.sierra@bbvinos.com",
            "BBVinos",
            "BBV1",
            "Las Parcelas 7950, Peñalolen",
            False,
            "BBVinosParser",
        ],
        [
            "rodrigo.curihuentro@bbvinos.com",
            "BBVinos",
            "BBV1",
            "Las Parcelas 7950, Peñalolen",
            False,
            "BBVinosParser",
        ],
        [
            "paula.sierra@bbvinos.com",
            "BBVinos",
            "BBV1",
            "Las Parcelas 7950, Peñalolen",
            False,
            "BBVinosParser",
        ],
        [
            "bodega-cl@prominent.com",
            "Prominent SpA",
            "PROM",
            "Diagonal Oriente 1755, Ñuñoa",
            False,
            None,
        ],
        [
            "Contreras.jhonny@prominent.com",
            "Prominent SpA",
            "PROM",
            "Diagonal Oriente 1755, Ñuñoa",
            False,
            None,
        ],
        [
            "Administracion-cl@prominent.com",
            "Prominent SpA",
            "PROM",
            "Diagonal Oriente 1755, Ñuñoa",
            False,
            None,
        ],
        [
            "Castillo.alex@promiment.com",
            "Prominent SpA",
            "PROM",
            "Diagonal Oriente 1755, Ñuñoa",
            False,
            None,
        ],
        [
            "juan@vdalcohuaz.cl",
            "Altas Tierras de Alcohuaz",
            "ATA1",
            "Presidente José Battle y Ordóñez 4835, Ñuñoa",
            False,
            None,
        ],
        [
            "matias@vdalcohuaz.cl",
            "Altas Tierras de Alcohuaz",
            "ATA1",
            "Presidente José Battle y Ordóñez 4835, Ñuñoa",
            False,
            None,
        ],
        [
            "pedidos@vdalcohuaz.cl",
            "Altas Tierras de Alcohuaz",
            "ATA1",
            "Presidente José Battle y Ordóñez 4835, Ñuñoa",
            False,
            None,
        ],
        [
            "javicaceres1831@gmail.com",
            "Bianca Rose Boutique",
            "BRB1",
            "Primo de Rivera 501, Maipu",
            False,
            None,
        ],
        [
            "Javieracastillocampos@hotmail.cl",
            "Arenas Cat",
            "ARC1",
            "Pasaje Padre Abdón Cifuentes 3312, Puente Alto",
            False,
            None,
        ],
        [
            "arenascatsantiago@gmail.com",
            "Arenas Cat",
            "ARC1",
            "Pasaje Padre Abdón Cifuentes 3312, Puente Alto",
            False,
            None,
        ],
        [
            "alan.saul.aravena@gmail.com",
            "Arenas Cat",
            "ARC1",
            "Pasaje Padre Abdón Cifuentes 3312, Puente Alto",
            False,
            None,
        ],
        [
            "cbarrientos@tslcargo.cl",
            "TSL Group Cargo SpA",
            "TSL1",
            "Armando Cortinez Oriente 945, Pudahuel",
            False,
            None,
        ],
        [
            "bgonzalez@tslcargo.cl",
            "TSL Group Cargo SpA",
            "TSL1",
            "Armando Cortinez Oriente 945, Pudahuel",
            False,
            None,
        ],
        [
            "cekuatro@gmail.com",
            "TSL Group Cargo SpA",
            "TSL1",
            "Armando Cortinez Oriente 945, Pudahuel",
            False,
            None,
        ],
        [
            "hgomez@axiomatix.cl",
            "Axiomatix SpA",
            "AXI1",
            "Fukui 6828, La Florida",
            False,
            None,
        ],
    ],
    columns=[
        "allowedEmail",
        "clientName",
        "codePrefix",
        "pickupAddress",
        "allowOverride",
        "customParser",
    ],
)
parsersDict = {
    "default": xls_import.xlsx_to_dispatches,
    "BBVinosParser": xls_import.bbvinos_xlsx_to_dispatches,
}


def check_if_allowed(filename):
    extension = filename.split(".")[-1]
    if extension.upper() in ALLOWED_FILES_EXTENSIONS:
        return True
    else:
        return False


def email_in_database(emailAddress: str) -> bool:
    allowedEmailList = [
        mail.upper() for mail in ALLOWED_CLIENTS_DF["allowedEmail"].to_list()
    ]
    if emailAddress.upper() in allowedEmailList:
        return True
    else:
        return False


def user_overrides(emailAddress):
    overridesList = (
        ALLOWED_CLIENTS_DF["allowedEmail"]
        .where(ALLOWED_CLIENTS_DF["allowOverride"] == True)
        .dropna()
        .tolist()
    )
    overridesList = [mail.upper() for mail in overridesList]
    if emailAddress.upper() in overridesList:
        return True
    else:
        return False


def main():
    LogicaAPI = BeetrackAPI(os.getenv("BEETRACK_APIKEY"), BASE_URL)
    MailOutbox = mail_handler.SMTPHandler(
        user=os.getenv("IMAP_USER"),
        passwd=os.getenv("IMAP_PASSWD"),
        server=os.getenv("SMTP_SERVER"),
        port=os.getenv("SMTP_PORT"),
    )
    MailInbox = mail_handler.Inbox(
        os.getenv("IMAP_USER"), os.getenv("IMAP_PASSWD"), os.getenv("IMAP_SERVER")
    )

    fetchedEmails = MailInbox.check_inbox()
    if len(fetchedEmails) == 0:
        timestamp = time.strftime("%H:%M:%S")
        print(f"[{timestamp}] No new mail found. Exiting.", file=sys.stdout)
        MailInbox.logout()
        return 0
    for email in fetchedEmails:
        if email_in_database(email._from):
            timestamp = time.strftime("%H:%M:%S")
            print(
                f"[{timestamp}] MailID {email.id} found from allowed client: {email._from}"
            )
        else:
            continue
        if any([check_if_allowed(att) for att in email.attachments]):
            print(f"MailID {email.id} contains a valid file", file=sys.stdout)
        else:
            print(f"MailID {email.id} doesn't contain any valid file. Skipping.")
            email.mark_read()
            continue
        clientName = (
            ALLOWED_CLIENTS_DF["clientName"]
            .where(ALLOWED_CLIENTS_DF["allowedEmail"] == email._from)
            .dropna()
            .iloc[0]
        )
        pickupAddress = (
            ALLOWED_CLIENTS_DF["pickupAddress"]
            .where(ALLOWED_CLIENTS_DF["allowedEmail"] == email._from)
            .dropna()
            .iloc[0]
        )
        if user_overrides(email._from):
            print(
                "[{timestamp}] Sender can override data. Reading mail body".format(
                    timestamp=time.strftime("%H:%M:%S")
                )
            )
            overridesDict = email.read_overrides(ALLOWED_CLIENTS_DF)
            if overridesDict["Cliente"]:
                clientName = overridesDict["Cliente"]
                print(
                    "[{timestamp}] Sender overrided client name: {clientName}".format(
                        timestamp=time.strftime("%H:%M:%S"), clientName=clientName
                    )
                )
            if overridesDict["pickup_address"]:
                pickupAddress = overridesDict["pickup_address"]
                print(
                    "[{timestamp}] Sender overrided pickup address: {pickupAddress}".format(
                        timestamp=time.strftime("%H:%M:%S"), pickupAddress=pickupAddress
                    )
                )
        timestamp = time.strftime("%H:%M:%S")
        reports = []
        for attachment in email.attachments:
            if not check_if_allowed(attachment):
                continue
            else:
                print(
                    f"[{timestamp}] Will scan file {os.path.basename(attachment)} for dispatches for client {clientName} to be picked up from {pickupAddress}",
                    file=sys.stdout,
                )
            # Verificar si hay que usar un parser Custom
            customParserQuery = (
                ALLOWED_CLIENTS_DF["customParser"]
                .where(ALLOWED_CLIENTS_DF["allowedEmail"] == email._from)
                .dropna()
            )
            if len(customParserQuery) > 0:
                chosenParser = parsersDict[customParserQuery[0]]
            else:
                chosenParser = parsersDict["default"]
            foundDispatchesData, warnings = chosenParser(
                attachment, clientName, pickupAddress
            )
            reportData = {
                "filename": os.path.basename(attachment),
                "general_issues": warnings,
                "dispatches": foundDispatchesData,
            }
            reports.append(reportData)
            if len(foundDispatchesData) == 0:
                print(
                    "[{timestamp}] Could not parse any dispatches in file {filename}:".format(
                        timestamp=time.strftime("%H:%M:%S"),
                        filename=os.path.basename(attachment),
                    ),
                    file=sys.stderr,
                )
                print(warnings, sep="\n", file=sys.stderr)
            elif len(warnings) > 0:
                print(
                    "[{timestamp}] There were errors importing dispatches from file {filename}:".format(
                        timestamp=time.strftime("%H:%M:%S"),
                        filename=os.path.basename(attachment),
                    ),
                    file=sys.stderr,
                )
                print(warnings, sep="\n", file=sys.stderr)
            else:
                pass
            for newDispatch, errorCode, warnings in foundDispatchesData:
                if errorCode == 2:
                    print(
                        "[{timestamp}] Failed to import dispatch {dispatchCode}:".format(
                            timestamp=time.strftime("%H:%M:%S"),
                            dispatchCode=newDispatch,
                        ),
                        file=sys.stderr,
                    )
                    print(*warnings, sep="\n", file=sys.stderr)
                    continue
                else:
                    response = LogicaAPI.create_dispatch(newDispatch.dump_dict())
                    if os.getenv("DEBUG"):
                        with open(
                            "{id}.json".format(id=newDispatch.id), "w"
                        ) as jsonFile:
                            json.dump(
                                newDispatch.dump_dict(),
                                jsonFile,
                                indent=4,
                                ensure_ascii=False,
                            )
                    print(response)
                if errorCode == 1:
                    print(
                        "[{timestamp}] Dispatch {dispatchCode} was imported with issues:".format(
                            timestamp=time.strftime("%H:%M:%S"),
                            dispatchCode=newDispatch.id,
                        ),
                        file=sys.stderr,
                    )
                    print(*warnings, sep="\n", file=sys.stderr)
                elif errorCode == 0:
                    print(
                        "[{timestamp}] Dispatch {dispatchCode} was imported successfully".format(
                            timestamp=time.strftime("%H:%M:%S"),
                            dispatchCode=newDispatch.id,
                        ),
                        file=sys.stdout,
                    )
                else:
                    print("CRITICAL: Unknown error code.")
        print(
            "[{timestamp}] Sending transactional email to {recipient}".format(
                timestamp=time.strftime("%H:%M:%S"), recipient=email._from
            ),
        )
        mail_handler.send_confirmation_mail(
            reports,
            _from=os.getenv("IMAP_USER"),
            to=email._from,
            subject=email.subject,
            outboxHandler=MailOutbox,
            replyingTo=email.emailObject,
        )
        email.mark_read()
    MailInbox.logout()
    print(
        "[{timestamp}] Checked all mails. Going to sleep.".format(
            timestamp=time.strftime("%H:%M:%S")
        )
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
