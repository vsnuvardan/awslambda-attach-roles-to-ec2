import json
class Status:

    def __init__(self,instanceId,region,status):
        self.instanceId = instanceId
        self.region = region
        self.status = status
    def toJSON(self):
        return json.dumps(self, default=lambda o: o.__dict__,
                          sort_keys=True, indent=4)
