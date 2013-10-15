import active_directory, os, _winreg, win32com.client

from jinja2 import Environment, FileSystemLoader

def generate_signatures(signature_path,ad_username):
    env = Environment(loader=FileSystemLoader(r'\\domain\netlogon\signatures\templates')) # signature template location

    #All templates
    html_template = env.get_template('signature_template.html')
    rtf_template = env.get_template('signature_template.rtf')
    plaintext_template = env.get_template('signature_template.txt')

    #Get AD fields
    user = active_directory.find_user(ad_username)
    displayName= user.displayName
    title = user.title
    company = user.company
    homePhone = None
    if user.homePhone:
        homePhone = user.homePhone
    telephoneNumber = user.telephoneNumber
    fax = user.facsimileTelephoneNumber
    email = user.mail
    website = user.wWWHomePage
    streetAddress = user.streetAddress
    city = user.l
    state = user.st
    zip = user.postalCode
    mobile = user.mobile
    #render HTML
    html_rendering = html_template.render(displayName = displayName,title = title,company = company,homePhone = homePhone,
                                telephoneNumber = telephoneNumber,fax = fax,email = email,website = website,
                                streetAddress = streetAddress,city = city,state = state, zip = zip, mobile = mobile)


    with open(signature_path + r"\default.htm", "wb") as f_html:
        f_html.write(html_rendering)
    f_html.close()

    #render RTF
    rtf_rendering = rtf_template.render(displayName = displayName,title = title,company = company,homePhone = homePhone,
                                telephoneNumber = telephoneNumber,fax = fax,email = email,website = website,
                                streetAddress = streetAddress,city = city,state = state, zip = zip, mobile = mobile)

    with open(signature_path + r"\default.rtf", "wb") as f_rtf:
        f_rtf.write(rtf_rendering)
    f_rtf.close()

    #render TXT
    txt_rendering = plaintext_template.render(displayName = displayName,title = title,company = company,homePhone = homePhone,
                                telephoneNumber = telephoneNumber,fax = fax,email = email,website = website,
                                streetAddress = streetAddress,city = city,state = state, zip = zip, mobile = mobile)
    txt_rendering = txt_rendering.replace('\n','\r\n')
    print txt_rendering
    with open(signature_path + r"\default.txt", "wb") as f_txt:
        f_txt.write(txt_rendering)
    f_txt.close()

def get_env_variables():
    username = os.environ.get("USERNAME") # local environmental variable
    appdata = os.environ.get("APPDATA") # local environmental variable
    print appdata
    signature_dir = r'' + appdata + r"\microsoft\signatures"
    if os.path.exists(signature_dir):
        return signature_dir, username
    else:
        os.makedirs(signature_dir)
        return signature_dir, username

def set_default():

    try:
        outlook_2013_key = _winreg.OpenKey(_winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\15.0\Common\MailSettings", 0, _winreg.KEY_ALL_ACCESS)
        _winreg.SetValueEx(outlook_2013_key, "NewSignature", 0, _winreg.REG_SZ, "default" )
        _winreg.SetValueEx(outlook_2013_key, "ReplySignature", 0, _winreg.REG_SZ, "default" )

        # set in outlook profile
        outlook_2013_base_key = _winreg.OpenKey(_winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\15.0\Outlook\Profiles", 0, _winreg.KEY_ALL_ACCESS)
        default_profile_2013_tup = _winreg.QueryValueEx(outlook_2013_base_key,'DefaultProfile')
        default_profile_2013 = default_profile_2013_tup[0]
        print default_profile_2013
        outlook_2013_profile_key = _winreg.OpenKey(_winreg.HKEY_CURRENT_USER,
                                                   "Software\\Microsoft\\Office\\15.0\\Outlook\\Profiles\\" + default_profile_2013 + "\\9375CFF0413111d3B88A00104B2A6676", 0, _winreg.KEY_ALL_ACCESS)
        for i in range(0, 10):
            try:
                outlook_2013_sub_key_name = _winreg.EnumKey(outlook_2013_profile_key,i)
                print outlook_2013_sub_key_name, "sub_key_name"
                outlook_2013_sub_key = _winreg.OpenKey(outlook_2013_profile_key, outlook_2013_sub_key_name, 0, _winreg.KEY_ALL_ACCESS)
                _winreg.SetValueEx(outlook_2013_sub_key, "New Signature", 0, _winreg.REG_SZ, "default" )
                _winreg.SetValueEx(outlook_2013_sub_key, "Reply-Forward Signature", 0, _winreg.REG_SZ, "default" )
            except:
                pass

    except:
        print('no 2013 found')


    try:
        outlook_2010_key = _winreg.OpenKey(_winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\14.0\Common\MailSettings", 0, _winreg.KEY_ALL_ACCESS)
        _winreg.SetValueEx(outlook_2010_key, "NewSignature", 0, _winreg.REG_SZ, "default" )
        _winreg.SetValueEx(outlook_2010_key, "ReplySignature", 0, _winreg.REG_SZ, "default" )
    except:
        print('no 2010 found')
    try:
        outlook_2007_key = _winreg.OpenKey(_winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\12.0\Common\MailSettings", 0, _winreg.KEY_ALL_ACCESS)
        _winreg.SetValueEx(outlook_2007_key, "NewSignature", 0, _winreg.REG_SZ, "default" )
        _winreg.SetValueEx(outlook_2007_key, "ReplySignature", 0, _winreg.REG_SZ, "default" )
    except:
        print('no 2007 found')
    try:

        outlook_2003_base_key = _winreg.OpenKey(_winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles", 0, _winreg.KEY_ALL_ACCESS)
        default_profile_tup = _winreg.QueryValueEx(outlook_2003_base_key,'DefaultProfile')
        default_profile = default_profile_tup[0]
        print default_profile
        outlook_2003_profile_key = _winreg.OpenKey(_winreg.HKEY_CURRENT_USER,
                                                   "Software\\Microsoft\\Windows NT\\CurrentVersion\\Windows Messaging Subsystem\\Profiles\\" + default_profile + "\\9375CFF0413111d3B88A00104B2A6676", 0, _winreg.KEY_ALL_ACCESS)
        for i in range(0, 10):
            try:
                outlook_2003_sub_key_name = _winreg.EnumKey(outlook_2003_profile_key,i)
                print outlook_2003_sub_key_name, "sub_key_name"
                outlook_2003_sub_key = _winreg.OpenKey(outlook_2003_profile_key, outlook_2003_sub_key_name, 0, _winreg.KEY_ALL_ACCESS)
                _winreg.SetValueEx(outlook_2003_sub_key, "New Signature", 0, _winreg.REG_SZ, "default" )
                _winreg.SetValueEx(outlook_2003_sub_key, "Reply-Forward Signature", 0, _winreg.REG_SZ, "default" )
            except:
                pass
    except:
        print('no 2003 found')




#generate_signatures('dboudwin')


if __name__ == "__main__":
    signature_path, ad_username = get_env_variables()
    generate_signatures(signature_path, ad_username)
    set_default()