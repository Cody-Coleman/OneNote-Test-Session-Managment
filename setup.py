from distutils.core import setup
import py2exe


setup( name='runreport',
		console=['runreport.py'],
		options={"py2exe":{
						"dll_excludes":[ "mswsock.dll", "MSWSOCK.dll", "powrprof.dll", 'w9xpopen.exe' ],
						"bundle_files": 1
						}
				},
		zipfile=None,
		)
