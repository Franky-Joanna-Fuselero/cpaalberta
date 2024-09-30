from collections import OrderedDict
import re
import string
from scrapy import Spider, Request, Selector


class cpaalberta(Spider):
    name = 'cpaalberta'

    custom_settings = {
        'FEED_EXPORTERS': {'xlsx': 'scrapy_xlsx.XlsxItemExporter'},
        'FEED_FORMAT': 'xlsx',
        'FEED_URI': 'cpaalberta_data.xlsx',
        'DOWNLOAD_DELAY': 0.2,
    }

    url = "https://services.cpaalberta.ca/VerifyEntity/Firms/ShowFirms"

    headers = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:128.0) Gecko/20100101 Firefox/128.0',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,'
                  'image/png,image/svg+xml,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.5',
        'Accept-Encoding': 'gzip, deflate, br, zstd',
        'Content-Type': 'application/x-www-form-urlencoded',
        'Origin': 'https://services.cpaalberta.ca',
        'Connection': 'keep-alive',
        'Upgrade-Insecure-Requests': '1',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-User': '?1',
        'Priority': 'u=0, i',
        'TE': 'trailers'
    }

    seen = []

    def start_requests(self):
        for alpha in string.ascii_lowercase:
            payload = f'firmname={alpha}&city=&Verify=Verify'
            meta = {'dont_merge_cookies': True}
            yield Request(self.url, self.parse_listing, meta=meta,
                          headers=self.headers, body=payload, method='POST')

    def parse_listing(self, response):
        row_ids = response.css('#datatable a::attr(onclick)').re('\((.+)\)')
        for row_id in row_ids:
            url = "https://services.cpaalberta.ca/VerifyEntity/Firms/ShowFirmDetails"
            payload = f'Clientid={row_id}'
            yield Request(url, self.parse_detail, meta=response.meta,
                          headers=self.headers, body=payload, method='POST')

    def parse_detail(self, response):
        item = OrderedDict()
        item['Firm Name'] = response.css('td:contains("Firm Name:") + td::text').get('').strip()
        item['Business City'] = response.css('td:contains("Business City:") + td::text').get('').strip()
        item['Standing'] = response.css('td:contains("Standing:") + td::text').get('').strip()
        item['Registered Service Areas'] = '\n'.join(self.get_set_text(response.css('td:contains("Registered Service Areas") + td + td ul')))
        item['Conditions/Restrictions'] = response.css('td:contains("Conditions/Restrictions") + td + td::text').get('').strip()
        yield item

    def get_set_text(self, selector, dont_skip=None):
        dont_skip = dont_skip or []
        assert isinstance(dont_skip, list), "'dont_skip' must be a 'list' or None type"

        required_tags = ['a', 'i', 'u', 'strong', 'b', 'em', 'span', 'sup', 'sub', 'font']
        required_tags.extend(dont_skip)

        results = []
        for text in selector.getall():
            for tag in required_tags:
                text = re.sub(r'<\s*%s>' % tag, '', text)
                text = re.sub(r'</\s*%s>' % tag, '', text)
                text = re.sub(r'<\s*%s[^\w][^>]*>' % tag, '', text)
                text = re.sub(r'</\s*%s[^\w]\s*>' % tag, '', text)

            text = text.replace('\r\n', ' ')
            text = re.sub(r'<!--.*?-->', '', text, re.S)
            sel = Selector(text=text)

            # extract all texts except tabular texts
            all_texts = sel.xpath(''.join([
                'descendant::text()/parent::*[name()!="td"]',
                '[name()!="script"][name()!="style"]/text()'
            ])).getall()
            all_texts = [x.strip() for x in all_texts]
            results += all_texts

        results = list(filter(None, results))
        return results
