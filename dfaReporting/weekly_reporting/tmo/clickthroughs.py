def strip_clickthroughs(data):

    data['Click-through URL'] = data['Click-through URL'].str.replace('http://analytics.bluekai.com/site/', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('%3F%3DADV_DS_%epid!_%eaid!_%ecid!_%eadv!', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('15991\?phint', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('http://15991\?phint', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('event%3Dclick&phint', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('aid%3D%eadv!&phint', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('pid%3D%epid!&phint', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('cid%3D%ebuy!&phint', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('crid%3D%ecid!&done', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('pid%3D%25epid!&phint', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('%3D%epid!_%eaid!_%ecid!_%eadv!', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('%26csdids', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('DADV_DS_ADDDVL4Q_EMUL7Y9E1YA4116', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('DWTR_DD_DDRDSPLYPR_JM2694TSP3U5895', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('DWTR_DD_DDRFCBK_RQLMKXRCUQZ1042', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('%3Fcmpid%3', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('b/refmh_', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace(
        '%3Fcm_mmc%3DPurchase-_-Display-_-Revere-_-Revere', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace(
        '%3Fcm_mmc%3DDisplay-_-Purchase-_-GM-_-Tablet_Base', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace(
        '%3Fcmpid%3DWTR_DD_DDRFCBK_RQLMKXRCUQZ1042%26csdids%3D%epid!_%eaid!_%ecid!_%eadv!', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace(
        '%3F%26csdids%3DADV_DS_%epid!_%eaid!_%ecid!_%eadv!', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace(
        '%3Fcm_mmc%3DPurchase-_-Display-_-Revere-_-Revere%26csdids%3D%epid!_%eaid!_%ecid!_%eadv!', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace(
        '%3Fcm_mmc%3DDisplay-_-Purchase-_-GM-_-Tablet_Base%26csdids%3D%epid!_%eaid!_%ecid!_%eadv!', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace(
        '%3Fcmpid%3DWTR_DD_DDRDSPLYPR_JM2694TSP3U5895%26csdids%3D%epid!_%eaid!_%ecid!_%eadv!', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('&csdids%epid!_%eaid!_%ecid!_%eadv!', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('=', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('%2F', '/')
    data['Click-through URL'] = data['Click-through URL'].str.replace('%3A', ':')
    data['Click-through URL'] = data['Click-through URL'].str.replace('%23', '#')
    data['Click-through URL'] = data['Click-through URL'].apply(lambda x: str(x).split('.html')[0])
    data['Click-through URL'] = data['Click-through URL'].apply(lambda x: str(x).split('?')[0])
    data['Click-through URL'] = data['Click-through URL'].apply(lambda x: str(x).split('%')[0])
    data['Click-through URL'] = data['Click-through URL'].apply(lambda x: str(x).split('_')[0])
    data['Click-through URL'] = data['Click-through URL'].str.replace('DWTR', '')

    return data