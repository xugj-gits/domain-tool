import pprint
import whois

domain = whois.whois('reanod.com')
pprint.pprint(domain)
print('name: ' + domain.domain_name)