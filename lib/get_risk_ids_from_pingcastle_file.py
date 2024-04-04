from argparse import ArgumentParser
from glob import glob
from bs4 import BeautifulSoup
from os.path import isfile
from lib.logging import log_call
from xml.etree import ElementTree

@log_call
def get_risk_ids_from_pingcastle_file(file_path:str):

        matching_html_files = glob(file_path)
        if matching_html_files:
            file_path = matching_html_files[0]
            if len(matching_html_files) == 1:
                print(f'PingCastle XML or HTML report found at "{file_path}"')
                print()
            elif len(matching_html_files) > 1:
                print(f'PingCastle XML or HTML report at "{file_path}" selected out of multiple options:')
                print("\n".join([f'-> {matching_html_file}' for matching_html_file in matching_html_files]))
                print()
        else:
            print(f'Impossible to find a PingCastle XML or HTML report at "{file_path}".')
            raise FileNotFoundError

        # IDS
        if ".xml" in file_path:
            try:
                risk_ids = get_risk_ids_from_pingcastle_xml_file(file_path)
            except FileNotFoundError as e:
                return None
        elif ".html" in file_path:
            try:
                risk_ids = get_risk_ids_from_pingcastle_html_file(file_path)
            except FileNotFoundError as e:
                return None
        else:
            print("The PingCastle file needs a \".xml\" or \".html\" extention")
            return None

        return risk_ids

@log_call
def get_risk_ids_from_pingcastle_html_file(file_path:str) -> list:

    matching_html_files = glob(file_path)
    if matching_html_files:
        file_path = matching_html_files[0]
        if len(matching_html_files) == 1:
            print(f'PingCastle HTML report found at "{file_path}"')
            print()
        elif len(matching_html_files) > 1:
            print(f'PingCastle HTML report at "{file_path}" selected out of multiple options:')
            print("\n".join([f'-> {matching_html_file}' for matching_html_file in matching_html_files]))
            print()
    else:
        print(f'Impossible to find a PingCastle HTML report at "{file_path}".')
        raise FileNotFoundError

    with open(file_path, 'r', encoding="utf-8") as file:
        html_content = file.read()
  
        soup = BeautifulSoup(html_content, 'html.parser')

        # <input type='hidden' name='json' value='{
        #   "generation":"2023-05-23 14:40:56Z",
        #   "version":"2.11.0.0",
        #   "users":4239,
        #   "computers":4477,
        #   "score":100,
        #   "anomaly":100,
        #   "staledobjects":100,
        #   "trust":100,
        #   "privilegedGroup":100,
        #   "maturityLevel":1,
        #   "rules":"A-Krbtgt,T-SIDFiltering,T-SIDHistoryUnknownDomain,P-DelegationGPOData,T-Inactive,S-OS-2003,P-Delegated,P-ServiceDomainAdmin,P-UnkownDelegation,S-OS-2008,A-CertTempCustomSubject,S-SIDHistory,S-PwdLastSet-90,S-DesEnabled,S-PwdNotRequired,P-ProtectedUsers,S-ADRegistrationSchema,P-AdminPwdTooOld,S-SMB-v1,S-OS-XP,P-Kerberoasting,S-OS-W10,A-AuditDC,P-ControlPathIndirectMany,S-OS-Win7,A-CertEnrollHttp,A-LAPS-Joined-Computers,S-PwdLastSet-DC,A-DCLdapSign,A-DCLdapsChannelBinding,S-WSUS-HTTP,P-UnconstrainedDelegation,S-Reversible,A-DnsZoneTransfert,A-PreWin2000Other,A-WeakRSARootCert2,S-PwdNeverExpires,P-LogonDenied,T-AlgsAES,A-DsHeuristicsLDAPSecurity,A-SHA1RootCert,A-DnsZoneAUCreateChild,A-NoServicePolicy,A-PreWin2000AuthenticatedUsers,P-DNSAdmin,A-NoNetSessionHardening,P-OperatorsEmpty,A-UnixPwd",
        #   "id":"8LxHr7jw1I+C4W2m13Q8LAnJgpmkek+h1G80D9icy3w="
        # }'>  
        input_element = soup.find('input', {'name': 'json'})
        
        if input_element is None:
            print("Impossible to find an \"input\" HTML element with the name \"json\"")
            return None
        
        value = input_element.get('value')
        
        if value is None:
            print("Impossible to find a \"value\" attribute in the \"input\" HTML element with the name \"json\"")
            return None
        
        start_patern = '"rules":"'
        start_patern_len = len(start_patern)
        end_patern = '"'

        start_index = value.find(start_patern)
        end_index = value.find(end_patern, start_index + start_patern_len)
        
        if start_index == -1 or end_index == -1:
            print("Impossible to find the sub-strings \"" + start_patern + "\" and \"" + end_patern + "\" in the \"value\" attribute of the \"input\" HTML element with the name \"json\"")
            return None
    
    substring = value[start_index + start_patern_len:end_index]
    risk_id_values = substring.split(',')
    
    print(f'Risk ids from the PingCastle HTML report at "{file_path}":\n{", ".join(risk_id_values)}')
    print()
    return risk_id_values

@log_call
def get_risk_ids_from_pingcastle_xml_file(file_path:str) -> list:

    matching_xml_files = glob(file_path)
    if matching_xml_files:
        file_path = matching_xml_files[0]
        if len(matching_xml_files) == 1:
            print(f'PingCastle XML report found at "{file_path}"')
            print()
        elif len(matching_xml_files) > 1:
            print(f'PingCastle XML report at "{file_path}" selected out of multiple options:')
            print("\n".join([f'-> {matching_xml_file}' for matching_xml_file in matching_xml_files]))
            print()
    else:
        print(f'Impossible to find the PingCastle XML report at "{file_path}".')
        raise FileNotFoundError

    # <HealthcheckData>
    #   <RiskRules>
    #     <HealthcheckRiskRule>
    #       <Points>50</Points>
    #       <Category>Anomalies</Category>
    #       <Model>GoldenTicket</Model>
    #   ->  <RiskId>A-Krbtgt</RiskId>
    #       <Rationale>Last change of the Kerberos password: 3108 day(s) ago</Rationale>
    #     </HealthcheckRiskRule>
    #     <HealthcheckRiskRule>
    #     ...
    #   <RiskRules>
    # </HealthcheckData>
    xml_tree = ElementTree.parse(file_path)
    root_element = xml_tree.getroot()
    risk_id_elements = root_element.findall('.//RiskId')
    risk_id_values = [risk_id_element.text for risk_id_element in risk_id_elements]
    print(f'Risk ids from the PingCastle XML report at "{file_path}":\n{", ".join(risk_id_values)}')
    print()
    return risk_id_values

@log_call
def request_file_path() -> str:
    file_path = None
    while not file_path:
        try:
            file_path = input("Path to the PingCastle HTML or XML file (Ctrl+C to quit) : ")
        except KeyboardInterrupt as e:
            raise KeyboardInterrupt
        if not isfile(file_path):
            print(f'Error : File not found at "{file_path}".')
            file_path = None
    return file_path

@log_call
def main():

    # ARGUMENTS
    parser = ArgumentParser(description='Parse a PingCastle HTML report and extract the list of the risks ID')
    parser.add_argument('-f', '--file', type=str, help='Path to a PingCastle HTML or XML file.')
    args = parser.parse_args()

    if hasattr(args, 'file') and args.file is not None:
        file_path = args.file 
    else:
        try:
            file_path = request_file_path()
        except KeyboardInterrupt:
            return
    # IDS
    return get_risk_ids_from_pingcastle_file(file_path)

if __name__ == '__main__':
    main()