from beetrack.beetrack_api import BeetrackAPI
from beetrack.beetrack_objects import Item, Dispatch

BASE_URL = "https://app.beetrack.com/api/external/v1"

# Import API KEY from file. This file should not be public
with open("APIKEY", "r") as APIKeyfile:
    API_KEY = APIKeyfile.read()

LogicaAPI = BeetrackAPI(API_KEY, BASE_URL)

testDispatch = LogicaAPI.get_dispatch("1004")
testDispatchObject = Dispatch(testDispatch["response"])

print(testDispatchObject.dump_dict())
print(testDispatchObject.dump_json())
