import filecmp
import os
import subprocess
import json


def print_conf_results(srv1,srv2,compare):
	return 0


def print_site_results(srv1,srv2,site,compare):
	
	flag = 0

	if compare.diff_files:
		print msgs['diff_found'] % (site,compare.diff_files)
		flag = 1
	if compare.left_only:
		print msgs['only_in'] % (site,srv1,compare.left_only)
		flag = 1
	if compare.right_only:
		print msgs['only_in'] % (site,srv2,compare.right_only)
		flag = 1
	if compare.funny_files:
		print msgs['comp_failed'] % (site,compare.funny_files)
		flag = 1
	
	if flag == 0:
		print msgs['ok'] % (site)




#COMPARE IIS CONFIGURATION
def compare_conf(srv1,srv2):
	str_cmd = '''%s -verb:sync \
	-source:webServer,computername=%s -disableLink:ContentExtension \
	-dest:webServer,computername=%s -whatif''' % (paths['web_deploy'], srv1, srv2)
	#print str_cmd
	run_compare = subprocess.Popen(str_cmd, shell=True, stdout = subprocess.PIPE)
	log_file = paths['conf_file'] % (srv1, srv2)
	with open(log_file, 'w') as f:
		for line in run_compare.stdout:
			f.write(line)
	print 'see configuration comparison at %s' % (log_file)



#COMPARE SITES
def compare_sites(srv1,srv2,site):
	dir1 = '\\\\%s\\%s\\%s' % (srv1, paths['root'], site_dirs[site])
	dir2 = '\\\\%s\\%s\\%s' % (srv2, paths['root'], site_dirs[site])

	dir1_exists = os.path.isdir(dir1)
	dir2_exists = os.path.isdir(dir2)

	try:
		compare = filecmp.dircmp(dir1,dir2)
		print_site_results(srv1,srv2,site,compare)
	except WindowsError:
		print msgs['dir_not_exist'] % dir1



#COMPARE SERVERS
def compare_servers(srv1,srv2):
	print '\n'
	print 'XML Compare Servers: %s vs. %s' % (srv1,srv2)
	for site, site_dir in site_dirs.items():
		compare_sites(srv1,srv2,site)

	print '\n'

	print 'Conf Compare Servers: %s vs. %s' % (srv1,srv2)
	compare_conf(srv1,srv2)



#COMPARE ENVIRONMENTS
def compare_envs(env1,env2):
	print 'Environments: %s vs. %s' % (env1,env2)
	for srv1_name, srv1_ip in iis_servers[env1].items():
		for srv2_name, srv2_ip in iis_servers[env2].items():
			compare_servers(iis_servers[env1][srv1_name],iis_servers[env2][srv2_name])
			print '\n'
	


#READ CONF
conf_file = 'L:/Scripts/compare_iis/properties.json'

with open(conf_file, 'rb') as fp:
	data = json.load(fp)
	paths = data["paths"]
	iis_servers = data["iis_servers"]
	site_dirs = data["site_dirs"]
	msgs = data["msgs"]
#END


compare_conf('172.16.21.241','172.16.0.12')
#compare_servers('172.16.21.241','172.16.0.12')
#compare_envs('qa','prod')