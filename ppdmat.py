#!/usr/bin/env python3
# PPDM assessment tool for Dell PowerProtect Data Manager - Github @ rjainoje
__author__ = "Raghava Jainoje"
__version__ = "1.0.2"
__email__ = " "
__date__ = "2023-09-26"

import argparse
from operator import index
from unicodedata import name
import requests
import urllib3
import sys
import json
import time
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
pd.options.mode.chained_assignment = None

writer = pd.ExcelWriter('ppdmdetails.xlsx', engine='xlsxwriter')
urllib3.disable_warnings()
summary_dict = {'PPDM SERVER DETAILS': ''}

def get_args():
    # Get command line args from the user
    parser = argparse.ArgumentParser(
        description='Script to gather PowerProtect Data Manager Information')
    parser.add_argument('-s', '--server', required=True,
                        action='store', help='PPDM DNS name or IP')
    parser.add_argument('-usr', '--user', required=False, action='store',
                        default='admin', help='User')
    parser.add_argument('-pwd', '--password', required=True, action='store',
                        help='Password')
    parser.add_argument('-rd', '--rptdays', required=False, action='store', default=30,
                        help='Report period')                    
    args = parser.parse_args()
    return args

def authenticate(ppdm, user, password, uri):
    # Login
    suffixurl = "/login"
    uri += suffixurl
    headers = {'Content-Type': 'application/json'}
    payload = '{"username": "%s", "password": "%s"}' % (user, password)
    try:
        response = requests.post(uri, data=payload, headers=headers, verify=False)
        response.raise_for_status()
    except requests.exceptions.ConnectionError as err:
        print('Error Connecting to {}: {}'.format(ppdm, err))
        sys.exit(1)
    except requests.exceptions.Timeout as err:
        print('Connection timed out {}: {}'.format(ppdm, err))
        sys.exit(1)
    except requests.exceptions.RequestException as err:
        print("The call {} {} failed with exception:{}".format(response.request.method, response.url, err))
        sys.exit(1)
    if (response.status_code != 200):
        raise Exception('Login failed for user: {}, code: {}, body: {}'.format(
            user, response.status_code, response.text))
    print('Logged in with user: {} to PPDM: {}'.format(user, ppdm))
    token = response.json()['access_token']
    return token

def get_appliance_config(uri, token):
    suffixurl = "/configurations"
    uri += suffixurl
    headers = {'Content-Type': 'application/json', 'Authorization': 'Bearer {}'.format(token)}
    pageSize = '10000'
    params = {'pageSize': pageSize}
    try:
        response = requests.get(uri, headers=headers, params=params, verify=False)
        response.raise_for_status()
    except requests.exceptions.RequestException as err:
        print("The call {}{} failed with exception:{}".format(response.request.method, response.url, err))
    if (response.status_code != 200):
        raise Exception('Failed to query {}, code: {}, body: {}'.format(uri, response.status_code, response.text))
    ac_df = pd.json_normalize(response.json()['content'], record_path=['networks'])
    return ac_df

def get_policies(uri, token):
    # Get all the policies
    suffixurl = "/protection-policies"
    uri += suffixurl
    headers = {'Content-Type': 'application/json', 'Authorization': 'Bearer {}'.format(token)}
    pageSize = '10000'
    filter = 'type eq "ACTIVE" and createdAt gt "2010-05-06T11:20:21.843Z"'
    params = {'filter': filter, 'pageSize': pageSize}
    try:
        response = requests.get(uri, headers=headers, params=params, verify=False)
        response.raise_for_status()
    except requests.exceptions.RequestException as err:
        print("The call {}{} failed with exception:{}".format(response.request.method, response.url, err))
    if (response.status_code != 200):
        raise Exception('Failed to query {}, code: {}, body: {}'.format(uri, response.status_code, response.text))
    FIELDS1 = ["name", "assetType", "type", "enabled", "encrypted", "dataConsistency", "summary.numberOfAssets", "summary.totalAssetCapacity", "summary.totalAssetProtectionCapacity", "summary.lastExecutionStatus"]
    df1 = pd.json_normalize(response.json()['content'])
    FIELDS2 = list(df1.keys())
    FIELDS = []
    for element in FIELDS1:
        if element in FIELDS2:
            FIELDS.append(element)
    po_df = df1[FIELDS]
    po_df.rename(columns={"name":'Name', "assetType":'AssetType', "type":'Type', "enabled":'Enabled', "encrypted":'Encrypted', "dataConsistency":'Data Consistency', "summary.numberOfAssets":'# of Assets', "summary.totalAssetCapacity":'TotalAssetCapacity(b)', "summary.totalAssetProtectionCapacity":'TotalAssetProtectionCapacity(b)', "summary.lastExecutionStatus":'Last Status'}, inplace=True)
    return po_df

def get_assets(uri, token):
    # Get all the Assets
    suffixurl = "/assets"
    uri += suffixurl
    headers = {'Content-Type': 'application/json', 'Authorization': 'Bearer {}'.format(token)}
    pageSize = '10000'
    filter = 'createdAt gt "2010-05-06T11:20:21.843Z"'
    params = {'filter': filter, 'pageSize': pageSize}
    try:
        response = requests.get(uri, headers=headers, params=params, verify=False)
        response.raise_for_status()
    except requests.exceptions.RequestException as err:
        print("The call {}{} failed with exception:{}".format(response.request.method, response.url, err))
    if (response.status_code != 200):
        raise Exception('Failed to query {}, code: {}, body: {}'.format(uri, response.status_code, response.text))
    FIELDS1 = ["id", "name", "type", "protectionStatus", "size", "subtype", "protectionPolicy.name", "protectionCapacity.size", "lastAvailableCopyTime", "details.k8s.inventorySourceName", "details.vm.guestOS", "details.vm.vcenterName", "details.vm.esxName", "details.database.clusterName"]
    df2 = pd.json_normalize(response.json()['content'])
    FIELDS2 = list(df2.keys())
    FIELDS = []
    for element in FIELDS1:
        if element in FIELDS2:
            FIELDS.append(element)
    as_df = df2[FIELDS]
    as_df.rename(columns={"name":'Name', "type":'Type', "protectionStatus":'Protection Status', "size":'Size', "subtype":'SubType', "protectionPolicy.name":'PolicyName', "protectionCapacity.size":'Protection Capacity(b)', "lastAvailableCopyTime":'LastBackupCopy', "details.k8s.inventorySourceName":'K8S Inv Source', "details.vm.guestOS":'VM Guest OS', "details.vm.vcenterName":'vCenterName', "details.vm.esxName":'ESX Name', "details.database.clusterName":'Database ClusterName'}, inplace=True)
    return as_df

def get_inv_src(uri, token):
    # Get all the Inventory Sources
    suffixurl = "/inventory-sources"
    uri += suffixurl
    headers = {'Content-Type': 'application/json', 'Authorization': 'Bearer {}'.format(token)}
    pageSize = '10000'
    params = {'pageSize': pageSize}
    try:
        response = requests.get(uri, headers=headers, params=params, verify=False)
        response.raise_for_status()
    except requests.exceptions.RequestException as err:
        print("The call {}{} failed with exception:{}".format(response.request.method, response.url, err))
    if (response.status_code != 200):
        raise Exception('Failed to query {}, code: {}, body: {}'.format(uri, response.status_code, response.text))
    FIELDS1 = ["name", "type", "version", "lastDiscoveryResult.status", "address"] 
    df3 = pd.json_normalize(response.json()['content'])
    FIELDS2 = list(df3.keys())
    FIELDS = []
    for element in FIELDS1:
        if element in FIELDS2:
            FIELDS.append(element)
    is_df = df3[FIELDS]
    return is_df

def get_storage(uri, token):
    # Get all the Storage Systems
    suffixurl = "/storage-systems"
    uri += suffixurl
    headers = {'Content-Type': 'application/json', 'Authorization': 'Bearer {}'.format(token)}
    pageSize = '10000'
    params = {'pageSize': pageSize}
    try:
        response = requests.get(uri, headers=headers, params=params, verify=False)
        response.raise_for_status()
    except requests.exceptions.RequestException as err:
        print("The call {}{} failed with exception:{}".format(response.request.method, response.url, err))
    if (response.status_code != 200):
        raise Exception('Failed to query {}, code: {}, body: {}'.format(uri, response.status_code, response.text))
    FIELDS1 = ["name", "type", "details.dataDomain.totalSize", "details.dataDomain.totalUsed", "capacityUtilization", "details.dataDomain.compressionFactor", "lastDiscoveryStatus", "lastDiscovered", "readiness", "details.dataDomain.version", "details.dataDomain.model", "details.dataDomain.serialNumber"] 
    df4 = pd.json_normalize(response.json()['content'])
    FIELDS2 = list(df4.keys())
    FIELDS = []
    for element in FIELDS1:
        if element in FIELDS2:
            FIELDS.append(element)
    st_df = df4[FIELDS]
    return st_df

def get_protection_eng(uri, token):
    # Get all the Protection Engines
    suffixurl = "/protection-engines"
    uri += suffixurl
    headers = {'Content-Type': 'application/json', 'Authorization': 'Bearer {}'.format(token)}
    pageSize = '10000'
    params = {'pageSize': pageSize}
    try:
        response = requests.get(uri, headers=headers, params=params, verify=False)
        response.raise_for_status()
    except requests.exceptions.RequestException as err:
        print("The call {}{} failed with exception:{}".format(response.request.method, response.url, err))
    if (response.status_code != 200):
        raise Exception('Failed to query {}, code: {}, body: {}'.format(uri, response.status_code, response.text))
    pe_df = pd.json_normalize(response.json()['content'])
    return pe_df

def get_app_agents(uri, token):
    # Get all the App Agents
    suffixurl = "/protection-engines"
    uri += suffixurl
    headers = {'Content-Type': 'application/json', 'Authorization': 'Bearer {}'.format(token)}
    pageSize = '10000'
    params = {'pageSize': pageSize}
    try:
        response = requests.get(uri, headers=headers, params=params, verify=False)
        response.raise_for_status()
    except requests.exceptions.RequestException as err:
        print("The call {}{} failed with exception:{}".format(response.request.method, response.url, err))
    if (response.status_code != 200):
        raise Exception('Failed to query {}, code: {}, body: {}'.format(uri, response.status_code, response.text))
    df7 = pd.json_normalize(response.json()['content'])
    print ("Written App agents information to ppdmdetails.xls")
    return df7

def get_activities(uri, token, window):
    # Get all the Activities
    suffixurl = "/activities"
    uri += suffixurl
    headers = {'Content-Type': 'application/json', 'Authorization': 'Bearer {}'.format(token)}
    filter = 'category eq "PROTECT" and classType in ("TASK") and state in ("COMPLETED") and createTime gt "{}"'.format(window)
    orderby = 'createTime DESC'
    pageSize = '10000'
    params = {'filter': filter, 'orderby': orderby, 'pageSize': pageSize}
    try:
        response = requests.get(uri, headers=headers, params=params, verify=False)
        response.raise_for_status()
    except requests.exceptions.RequestException as err:
        print("The call {}{} failed with exception:{}".format(response.request.method, response.url, err))
    if (response.status_code != 200):
        raise Exception('Failed to query {}, code: {}, body: {}'.format(uri, response.status_code, response.text))
    FIELDS1 = ["protectionPolicy.name", "asset.name", "category", "date", "createTime", "updateTime", "duration", "state", "result.status", "name", "host.name", "stats.assetSizeInBytes", "stats.bytesTransferred" , "stats.postCompBytes", "stats.dedupeRatio", "stats.reductionPercentage"]
    df8 = pd.json_normalize(response.json()['content'])
    # df8['date'] = pd.to_datetime(df8['createTime']).dt.date
    FIELDS2 = list(df8.keys())
    FIELDS = []
    for element in FIELDS1:
        if element in FIELDS2:
            FIELDS.append(element)
    ac_df = df8[FIELDS]
    ac_df.rename(columns={"protectionPolicy.name": 'Policy Name', "asset.name": "Asset Name", "category": "Category", "date": "Date", "duration": "Duration (sec)", "state": "State", "result.status": "Status", "name": "Task", "host.name": "Client Name", "stats.assetSizeInBytes": "Asset Size", "stats.bytesTransferred":"Data Transferred" , "stats.postCompBytes": "PostComp", "stats.dedupeRatio": "Dedupe Ratio", "stats.reductionPercentage": "Reduction %"}, inplace=True)
    ac_df = ac_df[ac_df["Policy Name"].notnull()]
    return ac_df

def get_jobgroups(uri, token, window):
    # Get all the JOB Groups
    suffixurl = "/activities"
    uri += suffixurl
    headers = {'Content-Type': 'application/json', 'Authorization': 'Bearer {}'.format(token)}
    filter = 'category eq "PROTECT" and classType in ("JOB_GROUP") and state in ("COMPLETED") and createdTime gt "{}"'.format(window)
    orderby = 'createTime DESC'
    pageSize = '10000'
    params = {'filter': filter, 'orderby': orderby, 'pageSize': pageSize}
    try:
        response = requests.get(uri, headers=headers, params=params, verify=False)
        response.raise_for_status()
    except requests.exceptions.RequestException as err:
        print("The call {}{} failed with exception:{}".format(response.request.method, response.url, err))
    if (response.status_code != 200):
        raise Exception('Failed to query {}, code: {}, body: {}'.format(uri, response.status_code, response.text))
    FIELDS1 = ["protectionPolicy.name", "protectionPolicy.type", "stats.numberOfAssets", "stats.numberOfProtectedAssets", "category", "subcategory", "classType", "startTime", "endTime", "duration", "stats.bytesTransferredThroughput", "state", "result.status", "stats.assetSizeInBytes", "stats.preCompBytes", "stats.postCompBytes", "stats.bytesTransferred", "stats.dedupeRatio", "stats.reductionPercentage"]
    df9 = pd.json_normalize(response.json()['content'])
    FIELDS2 = list(df9.keys())
    FIELDS = []
    for element in FIELDS1:
        if element in FIELDS2:
            FIELDS.append(element)
    jg_df = df9[FIELDS]
    jg_df['startTime'] = pd.to_datetime(jg_df['startTime']).dt.strftime('%Y-%m-%d %r')
    jg_df['endTime'] = pd.to_datetime(jg_df['endTime']).dt.strftime('%Y-%m-%d %r')
    jg_df.rename(columns={"protectionPolicy.name":'Policy Name', "protectionPolicy.type":'Policy Type', "stats.numberOfAssets":'# of Assets', "stats.numberOfProtectedAssets":'# of Protected Assets', "category":'Category', "subcategory":'SubCategory', "classType":'JobType', "duration":'Duration(sec)', "stats.bytesTransferredThroughput":'Throughput(bytes)', "result.status":'Status', "stats.assetSizeInBytes":'Asset Size(b)', "stats.preCompBytes":'PreComp(b)', "stats.postCompBytes":'PostComp(b)', "stats.bytesTransferred":'Bytes Transferred(b)', "stats.dedupeRatio":'Dedupe Ratio', "stats.reductionPercentage":'Reduction %'}, inplace=True)
    return jg_df


def get_ddmtrees(uri, token):
    # Get all the DD MTrees
    suffixurl = "/datadomain-mtrees"
    uri += suffixurl
    headers = {'Content-Type': 'application/json', 'Authorization': 'Bearer {}'.format(token)}
    pageSize = '10000'
    params = {'pageSize': pageSize}
    try:
        response = requests.get(uri, headers=headers, params=params, verify=False)
        response.raise_for_status()
    except requests.exceptions.RequestException as err:
        print("The call {}{} failed with exception:{}".format(response.request.method, response.url, err))
    if (response.status_code != 200):
        raise Exception('Failed to query {}, code: {}, body: {}'.format(uri, response.status_code, response.text))
    FIELDS1 = ["name", "type", "lastUpdated", "totalCapacityInBytes", "availableCapacityInBytes", "attributes.dayPreComp", "attributes.dayPostComp", "attributes.dayCompressionFactor", "attributes.usedLogicalCapacity", "attributes.serialNo", "_embedded.storageSystem.name", "retentionLockStatus", "retentionLockMode", "replicationTargets", "replicationSources", "createdAt", "attributes.groupId", "attributes.user"]
    df10 = pd.json_normalize(response.json()['content'])
    FIELDS2 = list(df10.keys())
    FIELDS = []
    for element in FIELDS1:
        if element in FIELDS2:
            FIELDS.append(element)
    dd_df = df10[FIELDS]
    return dd_df

def get_license(uri, token):
    # Get all the Licenses
    suffixurl = "/licenses"
    uri += suffixurl
    headers = {'Content-Type': 'application/json', 'Authorization': 'Bearer {}'.format(token)}
    pageSize = '10000'
    params = {'pageSize': pageSize}
    try:
        response = requests.get(uri, headers=headers, params=params, verify=False)
        response.raise_for_status()
    except requests.exceptions.RequestException as err:
        print("The call {}{} failed with exception:{}".format(response.request.method, response.url, err))
    if (response.status_code != 200):
        raise Exception('Failed to query {}, code: {}, body: {}'.format(uri, response.status_code, response.text))
    li_df = pd.json_normalize(response.json()['content'], record_path=['licenseKeys'])
    licdict = li_df.to_dict('records')
    return licdict

def get_srvdr(uri, token):
    # Get PPDM Server DR Copies
    suffixurl = "/server-disaster-recovery-backups"
    uri += suffixurl
    headers = {'Content-Type': 'application/json', 'Authorization': 'Bearer {}'.format(token)}
    pageSize = '10000'
    params = {'pageSize': pageSize}
    try:
        response = requests.get(uri, headers=headers, params=params, verify=False)
        response.raise_for_status()
    except requests.exceptions.RequestException as err:
        print("The call {}{} failed with exception:{}".format(response.request.method, response.url, err))
    if (response.status_code != 200):
        raise Exception('Failed to query {}, code: {}, body: {}'.format(uri, response.status_code, response.text))
    srvdr_df = pd.json_normalize(response.json()['content'])
    FIELDS1 = ["hostname", "name", "version", "state", "creationTime", "backupConsistencyType", "components"]
    FIELDS2 = list(srvdr_df.keys())
    FIELDS = []
    for element in FIELDS1:
        if element in FIELDS2:
            FIELDS.append(element)
    srvdrb_df = srvdr_df[FIELDS]
    srvdict = srvdrb_df.to_dict('records')
    summary_dict['PPDM Hostname'] = srvdict[0]['hostname']
    summary_dict['PPDM Version'] = srvdict[0]['version']
    return srvdrb_df

def chartxls(activities):
    # Create a chart
    activities['Date'] = pd.to_datetime(activities['createTime']).dt.date
    act_chartdf = activities[["Date", "Asset Size", "PostComp", "Dedupe Ratio"]]
    act_chartdf["BackupSize-GB"] = round(act_chartdf["Asset Size"] / 1024 / 1024 / 1024, 2)
    act_chartdf["PostComp-GB"] = round(act_chartdf["PostComp"] / 1024 / 1024 / 1024, 2)
    act_chartdf["Date"] = act_chartdf["Date"]
    abc = act_chartdf.groupby('Date').agg({'BackupSize-GB':'sum', 'PostComp-GB': 'sum'})
    abc.to_excel(writer, sheet_name='Chart')
    workbook = writer.book
    worksheet = writer.sheets['Chart']
    chart = workbook.add_chart({'type': 'column'})
    max_row = len(abc) + 1
    for i in range(len(['Date', 'BackupSize-GB'])):
        col = i + 1
        chart.add_series({
        'name':       ['Chart', 0, col],
        'categories': ['Chart', 1, 0, max_row, 0],
        'values':     ['Chart', 1, col, max_row, col],
        'line':       {'width': 1.00},
        })
    chart.set_x_axis({'name': 'Date', 'date_axis': True})
    chart.set_y_axis({'name': 'Size(GB)', 'major_gridlines': {'visible': False}})
    chart.set_legend({'position': 'top'})
    chart.set_size({'width': 900, 'height': 576})
    worksheet.insert_chart('E2', chart)
    print ("Created column chat to ppdmreport.xls")

def summaryxls(assets, activities, jobgroups, ddmtrees, licinfo, rptdays):
    # Write summary to excel sheet named summary
    if licinfo[0]['featureName'] == "POWERPROTECT SW TRIAL":
        summary_dict['License Type'] = licinfo[0]['featureName']
        summary_dict['Expiry Date'] = licinfo[0]['endDate']
    else:
        summary_dict['License Type'] = licinfo[0]['featureName']
        summary_dict['Expiry Date'] = licinfo[0]['licenseType']
    summary_dict['ASSET SUMMARY'] = ''
    atype = assets.value_counts('Type')
    astatus = assets.value_counts('Protection Status')
    asize = assets['Size'].sum()
    fetb = assets['Protection Capacity(b)'].sum()
    summary_dict.update(atype.to_dict())
    summary_dict.update(astatus.to_dict())
    summary_dict['Total Assets Size (GB)'] = round(asize/1024/1024/1024, 2)
    summary_dict['Protection Size (GB) - FETB'] = round(fetb/1024/1024/1024, 2)
    act_status = activities.value_counts('Status')
    act_assetsize = activities['Asset Size'].sum()
    act_bytestrans = activities['Data Transferred'].sum()
    act_postcomp = activities['PostComp'].sum()
    act_lowdedupe = activities['Dedupe Ratio']
    dedupeless1 = act_lowdedupe[act_lowdedupe < 1].count()
    dedupeless3 = act_lowdedupe[act_lowdedupe < 3].count()
    dedupegt3 = act_lowdedupe[act_lowdedupe > 3].count()
    mtree_precomp = ddmtrees['attributes.dayPreComp'].astype(float).sum()
    mtree_postcomp = ddmtrees['attributes.dayPostComp'].astype(float).sum()
    summary_dict['ACTIVITIES SUMMARY - {} DAYS'.format(rptdays)] = ''
    summary_dict.update(act_status.to_dict())
    summary_dict['Backup Size (GB)'] = round(act_assetsize/1024/1024/1024, 2)
    summary_dict['Transferred Size (GB)'] = round(act_bytestrans/1024/1024/1024, 2)
    summary_dict['PostComp Size (GB)'] = round(act_postcomp/1024/1024/1024, 2)
    summary_dict['Compression (< 1x) Clients'] = dedupeless1
    summary_dict['Compression (< 3x) Clients'] = dedupeless3
    summary_dict['Compression (> 3x) Clients'] = dedupegt3
    jbstats = jobgroups['Throughput(bytes)']
    lessth1mb = jbstats[jbstats < 1000000].count()
    lessth5mb = jbstats[jbstats < 5000000].count()
    summary_dict['Backup Throughput (< 1MB) Clients'] = lessth1mb
    summary_dict['Backup Throughput (< 5MB) Clients'] = lessth5mb
    summary_dict['DATA DOMAIN SUMMARY - LAST DAY'] = ''
    summary_dict['PreComp (GB)'] = round(mtree_precomp/1024/1024/1024, 2)
    summary_dict['PostComp (GB)'] = round(mtree_postcomp/1024/1024/1024, 2)
    summdf = pd.DataFrame(list(summary_dict.items()), columns = ['Name','Value'])
    summdf.to_excel(writer, sheet_name='Summary', index=False)
    worksheet = writer.sheets['Summary']
    (max_row, max_col) = summdf.shape
    column_settings = [{'header': column} for column in summdf.columns]
    worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings, 'style': 'Table Style Medium 2'})
    worksheet.set_column(0, max_col - 1, 12)
    for column in summdf:
        column_length = max(summdf[column].astype(str).map(len).max(), len(column))
        col_idx = summdf.columns.get_loc(column)
        writer.sheets['Summary'].set_column(col_idx, col_idx, column_length)
    print ("Written Summary information to ppdmdetails.xls")

def outxls(df_dict):
    # Write output to excel
    for sheet, df in  df_dict.items():
        df.to_excel(writer, sheet_name = sheet, startrow=1, header=False, index=False)
        workbook = writer.book
        worksheet = writer.sheets[sheet]
        (max_row, max_col) = df.shape
        column_settings = [{'header': column} for column in df.columns]
        worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings, 'style': 'Table Style Medium 2'})
        worksheet.set_column(0, max_col - 1, 12)
        print ("Written '{}' information to ppdmdetails.xls".format(sheet))
    # writer.sheets['Summary'].activate()
    writer.close()

def logout(ppdm, user, uri, token):
    suffixurl = "/logout"
    uri += suffixurl
    headers = {'Content-Type': 'application/json', 'Authorization': 'Bearer {}'.format(token)}
    try:
        response = requests.post(uri, headers=headers, verify=False)
        response.raise_for_status()
    except requests.exceptions.RequestException as err:
            print("The call {} {} failed with exception:{}".format(response.request.method, response.url, err))
    if (response.status_code != 204):
        raise Exception('Logout failed for user: {}, code: {}, body: {}'.format(
            user, response.status_code, response.text))
    print('Logout for user: {} from PPDM: {}'.format(user, ppdm))


def main():
    port = "8443"
    apiendpoint = "/api/v2"
    args = get_args()
    ppdm, user, password, rptdays = args.server, args.user, args.password, args.rptdays
    uri = "https://{}:{}{}".format(ppdm, port, apiendpoint)
    token = authenticate(ppdm, user, password, uri)
    gettime = datetime.now() - timedelta(days = int(rptdays))
    window = gettime.strftime('%Y-%m-%dT%H:%M:%S.%fZ')
    appconfig = get_appliance_config(uri, token) 
    policies = get_policies(uri, token)
    assets = get_assets(uri, token)
    invsources = get_inv_src(uri, token)
    storage = get_storage(uri, token)
    protectioneng = get_protection_eng(uri, token)
    appagents = get_app_agents(uri, token)
    activities = get_activities(uri, token, window)
    jobgroups = get_jobgroups(uri, token, window)
    ddmtrees = get_ddmtrees(uri, token)
    licinfo = get_license(uri, token)
    try:
        srvdrinfo = get_srvdr(uri, token)
    except:
        srvdrinfo = pd.DataFrame()
    summaryxls(assets, activities, jobgroups, ddmtrees, licinfo, rptdays)
    try:
        chartxls(activities)
    except:
        pass
    df_dict = {'Activities': activities, 'JobGroups': jobgroups, 'Policies': policies, 'Assets': assets, 'InvSources': invsources, 'Storage': storage, 'DDStorageUnits': ddmtrees, 'ProtectionEngines': protectioneng, 'AppAgents': appagents, 'PPDMServer': appconfig, 'ServerDR': srvdrinfo}
    outxls(df_dict)
    print("All the data written to the file")
    logout(ppdm, user, uri, token)

if __name__ == "__main__":
    main()