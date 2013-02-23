#iis_sync

import csv,subprocess,sys
from datetime import datetime


paths = {
	'script'        : 'c:\\scripts\\py_sync_iis'
	'iis_sites'     : 'c:\\scripts\\py_sync_iis\\iis_sites.txt'	
	'web_deploy'    : '"c:\\Program Files\\IIS\\Microsoft Web Deploy V2\\msdeploy"'
}


now = datetime.now()
time_stamp = now.strftime('%Y_%m_%d-%H%M')
sites_synced = ''
sites_failed = ''


err = {
	'sync_ok'      : 'SYNC OK -go away',
	'sync_err'     : 'SYNC FAILED (see log in primary IIS server) for sites: ',
	'sync_unknown' : 'SYNC STATUS UNKNOWN, sync service running?',
	'index_error'  : 'sites file indexing error - please check that your iis sites file is indexed correctly'
}



with open(paths['iis_sites'], 'rb') as file_iis_sites:
	iis_sites_table = csv.reader(file_iis_sites, delimiter=',', quoting=csv.QUOTE_NONE)
	for iis_site in iis_sites_table:
		if iis_site[0].startswith(':'):
			try:
				iis_srv_src = iis_site[0][1:]
				iis_srv_dst = iis_site[1]
			except IndexError:
				print err['index_error']
			flag_sync_ok = 1
		elif not iis_site[0].startswith(';') \
			and not iis_site[0].startswith(' '):
			try:
				iis_site_id = iis_site[0]
				iis_site_name = iis_site[1]
				iis_site_port = iis_site[2]
			except IndexError:
				print err['index_error']

			path_iis_sync_log = '%s\SYNC.%s_%s_%s_%s.log' % (paths['script'],iis_srv_src,iis_srv_dst,iis_site_name,time_stamp)
			path_iis_err_log = '%s\ERR.%s_%s_%s_%s.log' % (paths['script'],iis_srv_src,iis_srv_dst,iis_site_name,time_stamp)
			str_sync = paths['web_deploy'] + ' -verb:sync' \
					+ ' -source:metakey=lm/w3svc/' + iis_site_id + ',computername=' + iis_srv_src \
			  		+ ' -dest:metakey=lm/w3svc/' + iis_site_id + ',computername=' + iis_srv_dst \
					+ ' -enableLink:AppPoolExtension'

			with open(path_iis_sync_log, 'w') as file_iis_sync_log, \
				open(path_iis_err_log, 'w') as file_iis_err_log:
				run_sync = subprocess.Popen(str_sync, shell=True, stdout = subprocess.PIPE)
				for line in run_sync.stdout:
					file_iis_sync_log.write(line)
					if not line.startswith('Info') \
						and not line.startswith('Total'):
						flag_sync_ok = 0
						file_iis_err_log.write(line)
				if flag_sync_ok == 1:
					file_iis_err_log.write(err['sync_ok'])
					sites_synced += iis_site_name + ' '
				elif flag_sync_ok == 0:
					sites_failed += iis_site_name + ' '