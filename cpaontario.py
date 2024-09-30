from collections import OrderedDict
from scrapy import Spider, Request


class cpaontario(Spider):
    name = 'cpaontario'

    custom_settings = {
        'FEED_EXPORTERS': {'xlsx': 'scrapy_xlsx.XlsxItemExporter'},
        'FEED_FORMAT': 'xlsx',
        'FEED_URI': 'cpaontario_data.xlsx',
    }

    url = "https://myportal.cpaontario.ca/s/sfsites/aura?r=6&aura.ApexAction.execute=1"

    payload = ("message={\"actions\":[{\"id\":\"132;a\",\"descriptor\":\"aura://ApexActionController/"
               "ACTION$execute\",\"callingDescriptor\":\"UNKNOWN\",\"params\":{\"namespace\":\"\","
               "\"classname\":\"FirmDirectoryController\",\"method\":\"getFirmsList\",\"params\":"
               "{\"pagenumber\":1,\"numberOfRecords\":20000,\"pageSize\":20000},\"cacheable\":"
               "false,\"isContinuation\":false}}]}&aura.context={\"mode\":\"PROD\",\"fwuid\""
               ":\"\",\"app\":\"siteforce:communityApp\",\"loaded\":{\"APPLICATION@markup"
               "://siteforce:communityApp\":\"vgD8vvaBHzgKYqb_JQjQdw\",\"MODULE@markup:"
               "//lightning:f6Controller\":\"\",\"COMPONENT@markup://instrumentation:"
               "o11ySecondaryLoader\":\"\"},\"dn\":[],\"globals\":{},\"uad\":false}&"
               "aura.pageURI=/s/firm-directory&aura.token=null")

    headers = {
        'accept': '*/*',
        'accept-language': 'en-US,en;q=0.9',
        'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'origin': 'https://myportal.cpaontario.ca',
        'priority': 'u=1, i',
        'referer': 'https://myportal.cpaontario.ca/s/firm-directory',
        'sec-ch-ua': '"Not/A)Brand";v="8", "Chromium";v="126", "Google Chrome";v="126"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"macOS"',
        'sec-fetch-dest': 'empty',
        'sec-fetch-mode': 'cors',
        'sec-fetch-site': 'same-origin',
        'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36'
                      '(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
        'x-sfdc-lds-endpoints': 'ApexActionController.execute:FirmDirectoryController.getFirmsList',
    }

    def start_requests(self):
        meta = {'dont_merge_cookies': True}
        yield Request(self.url, self.parse_listing, meta=meta,
                      headers=self.headers, body=self.payload, method='POST')

    def parse_listing(self, response):
        for row in response.json()['actions'][0]['returnValue']['returnValue']:
            item = OrderedDict()
            item['Firm Name'] = row['CPAO_Firm_Name_Dir__c']
            item['Business Street'] = row.get('BillingStreet', '')
            item['Business City'] = row.get('BillingCity', '')
            item['Business Province'] = row.get('BillingCity', '')
            yield item
