#
# Copyright (c) 2020 Pearce Software Solutions. All rights reserved.
#
import http.client
import json
import sys
import os

#
# from base64 import b64encode
#
class CommonRequest:
    def __init__(self, action, url, hdrs=None, data=None):

        self.ssl_flag = False
        self.action = action
        self.status = -1
        self.response = None
        self.headers = hdrs
        self.data = data
        self.responseData = None
        self.setUpParts(url)

        # go ahead and set up the connection
        try:
            if self.ssl_flag == True:
                self.connection = http.client.HTTPSConnection(self.dns)
            else:
                self.connection = http.client.HTTPConnection(self.dns)

        except Exception as e:
            print("Exception:" + str(e))
            print("> ", self)
            print("Unexpected error:" + str(sys.exc_info()[0]))
            # self.connection.close()
            return False

    def __delete__(self, instance):
        if hasattr(self, "connection"):
            self.connection.close()
        # print("deleting:",self)
        # self.Close() # just to be sure the connection has been closed and terminated

    def __str__(self):

        return (
            "Action: "
            + self.action
            + " SSL:"
            + str(self.ssl_flag)
            + " dns:"
            + self.dns
            + " url: "
            + self.url
            + " header: "
            + str(self.headers)
            + " data "
            + str(self.data)
            + " status: "
            + str(self.Status())
            + " response data:"
            + str(self.ResponseData())
        )

    def setUpParts(self, url):
        # "https://cloud.iexapis.com/v1/stock/"
        # "http://localhost.5984/db/_id
        # "https://www.yahoo.com"
        #
        parts = url.split("/", 3)
        #
        # expecting
        # 0 - http or https
        # 1 - empty
        # 2 dns[:port]
        # 3 everything else
        # print(len(parts), ":", parts)
        if len(parts) < 3:
            print("Error: Expecting more parts")
            return

        # anything else is not ssl
        if parts[0].lower() == "https:":
            self.ssl_flag = True

        self.dns = parts[2]

        if len(parts) == 4:
            self.url = "/" + parts[3]
        else:
            self.url = ""

        # print("setUpParts:", str(self))

    def Clear(self):
        self.url = None
        self.headers = None
        self.data = None
        self.response = None
        self.responseData = None

    def Request(self, action=None, url=None, hdrs=None, data=None):

        # can things change on the fly?
        if url != None:
            # can change url on the fly?
            orig_ssl_flag = self.ssl_flag
            self.setUpParts(url)
            if orig_ssl_flag != self.ssl_flag:
                print("Error: Cannot change SSL types.  Create a new object")
                # TODO: How to change ssl type on the fly
                return False

        if hdrs != None and isinstance(hdrs, dict):
            # headers must be JSON
            self.headers = hdrs

        if self.headers != None and not isinstance(self.headers, dict):
            print("Something is wrong with the headers")
            print(self)
            return

        if data != None:
            self.data = data

        if action != None:
            self.action = action

        if data != None and isinstance(data, dict):
            # print("WARNING: data in json dict")
            payload = json.dumps(data)
            payload.encode("utf-8")
        else:
            payload = self.data

        try:
            # print("starting request", url, self.headers, payload)
            if self.data == None:
                if self.headers == None:
                    self.connection.request(self.action, self.url)
                else:
                    self.connection.request(self.action, self.url, headers=self.headers)
            else:
                self.connection.request(
                    self.action, self.url, payload, headers=self.headers
                )

            self.response = self.connection.getresponse()

            if self.response == None:
                print("ERROR: None Response")
                self.status = -1
                return None

            self.responseData = self.response.read().decode()

            if self.response.status >= 200 and self.response.status < 300:
                return True
            else:
                return False

        except Exception as e:
            print("Exception:" + str(e))
            print(self)
            print("Unexpected error:" + str(sys.exc_info()[0]))
            # self.connection.close()
            return False

    def Status(self):
        if hasattr(self, "response"):
            if self.response != None and hasattr(self, "status"):
                return self.response.status
        return -1

    def ResponseData(self):
        if hasattr(self, "responseData"):
            return self.responseData
        return None

    def Close(self):
        if hasattr(self, "connection"):
            self.connection.close()
        self.action = ""
        self.dns = ""
        self.ssl_flag = False
        self.status = -1
        self.response = None
        self.headers = None
        self.data = None
        self.url = None
        self.responseData = None


"""
    url = (
            "https://cloud.iexapis.com/v1/stock/"
            + self.symbol.lower()
            + "/batch?types=quote,stats,news,dividends&range=1y&last=3&token="
            + self.token
    )
"""

if __name__ == "__main__":

    # conn = CommonRequest("cloud.iexapis.com",True)
    token = os.getenv("TOKEN", "junk")

    url = "https://cloud.iexapis.com/v1/stock/HD/batch?types=quote&token=" + token
    # print(url)
    conn = CommonRequest("GET", url)
    if conn == None:
        print("wtw?")
        sys.exit(-1)

    if conn.Request():  # "GET", url):
        print("success:", conn)
        # print(conn.ResponseData())
    else:
        print("error:", conn)
        # print(conn.Status())
        # print(conn.ResponseData())

    # test Closing
    conn.Close()
    # test deleting
    del conn

    conn = CommonRequest("GET", "http://localhost:5984")
    if conn.Request():  #
        print("success:", conn)
        # print(conn.ResponseData())
    else:
        print("error:", conn)
        # print(conn.Status())
        # print(conn.ResponseData())
