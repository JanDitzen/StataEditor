# -------------------------------------------------------------
# Imports and Constants
# -------------------------------------------------------------
import sublime, sublime_plugin
import Pywin32.setup, win32com.client, win32con, win32api
import os, tempfile, subprocess, re, urllib, json, random, time, calendar, winreg
from collections import defaultdict
# http://msdn.microsoft.com/en-us/library/windows/desktop/ms633548(v=vs.85).aspx

settings_file = "StataEditor.sublime-settings"
stata_debug = False

# -------------------------------------------------------------
# Classes
# -------------------------------------------------------------

class StataBuildCommand(sublime_plugin.WindowCommand):
	def run(self, **kwargs):
		getView().window().run_command("stata_execute", {"build":True, "Mode": kwargs["Mode"]})

class StataUpdateJsonCommand(sublime_plugin.TextCommand):
	"""Update the .json used in Stata dataset/varname autocompletions"""
	def run(self, edit):
		get_autocomplete_data(self.view, force_update=True, add_from_buffer=False, obtain_varnames=True)

class StataAutocompleteDtaCommand(sublime_plugin.TextCommand):
	def run(self, edit):
		datasets = get_autocomplete_data(self.view, add_from_buffer=True, obtain_varnames=False)
		if datasets is None:
			return
		self.suggestions = sorted( list(zip(*datasets))[1] ) # Tuple (fn, dta name)
		self.view.window().show_quick_panel(self.suggestions, self.insert_link) #, sublime.MONOSPACE_FONT)

	def insert_link(self, choice):
		if choice==-1:
			return
		link = '"' + self.suggestions[choice] + '"'
		self.view.run_command("stata_insert", {'link':link})

class StataAutocompleteVarCommand(sublime_plugin.TextCommand):
	def run(self, edit, menu='all', prev_choice=-1, filter_dta=None):

		# Three menus: normal ("all"), select one DTA only ("filter"), pick which dta to select ("dta")
		assert menu in ('all', 'filter', 'dta')
		self.menu = menu
		self.filter_dta = filter_dta

		# dtamap: dict of dta->varlist
		# datasets: list of dtas
		# varlist: dict of varlist -> datasets

		dtamap = get_autocomplete_data(self.view, add_from_buffer=True, obtain_varnames=True)
		if dtamap is None:
			return

		if menu=='all':
			varlist = defaultdict(list)
			for dta,variables in dtamap.items():
				for varname in variables:
					varlist[varname].append(dta)
			if not varlist: return
			self.suggestions = [['    ----> Select this to filter by dataset <----    ','']] + list( [v, ' '.join(d)] for v,d in varlist.items() )
		elif menu=='filter':
			varlist = dtamap[filter_dta]
			if not varlist: return
			self.suggestions = ['    ----> Variables in {} <----    '.format(filter_dta)] + sorted(varlist)
		else:
			self.datasets = dtamap.keys()
			if not self.datasets: return
			self.suggestions = [['    ----> Remove filter <----    ', '']] + sorted([d, ' '.join(v)] for d,v in dtamap.items())

		if prev_choice+1>=len(self.suggestions):
			prev_choice = -1
		sublime.set_timeout(lambda: self.view.window().show_quick_panel(self.suggestions, self.insert_link, selected_index=prev_choice+1), 1) #, flags=sublime.MONOSPACE_FONT)

	def insert_link(self, choice):
		# Lots of recursive calls; alternatively I could just have a while loop
		if choice==-1:
			return

		if choice==0:
			if self.menu=='all':
				self.run(None, menu='dta', prev_choice=0)
			else:
				self.run(None, menu='all')
			return

		if self.menu=='all':
			link = self.suggestions[choice][0] + ' '
		elif self.menu=='filter':
			link = self.suggestions[choice] + ' '
		else:
			link = self.suggestions[choice][0]
			self.run(None, menu='filter', filter_dta=link, prev_choice=0)
			return

		self.view.run_command("stata_insert", {'link':link})
		
		# Call again until the user presses Escape
		self.run(None, menu=self.menu, filter_dta=self.filter_dta, prev_choice=choice)

class StataInsert(sublime_plugin.TextCommand):
	def run(self, edit, link):
		startloc = self.view.sel()[-1].end()
		self.view.insert(edit, startloc, link)

class StataExecuteCommand(sublime_plugin.TextCommand):
	def run(self, edit, **args):
		all_text = ""
		len_sels = 0
		sels = self.view.sel()
		len_sels = 0
		for sel in sels:
			len_sels = len_sels + len(sel)

		if len_sels==0 or args.get("Build", False)==True:
			all_text = self.view.substr(self.view.find('(?s).*',0))
		else:
			self.view.run_command("expand_selection", {"to": "line"})
			for sel in sels:
				all_text = all_text + self.view.substr(sel)

		if all_text[-1] != "\n":
			all_text = all_text + "\n"

		dofile_path = os.path.join(tempfile.gettempdir(), 'st_stata_temp.tmp')

		this_file = open(dofile_path,'w')
		this_file.write(all_text)
		this_file.close()
		
		view = self.view.window().active_view()
		cwd = get_cwd(view)
		if cwd: StataAutomate("cd " + cwd)
		
		StataAutomate(str(args["Mode"]) + " " + dofile_path)

class StataHelpExternal(sublime_plugin.TextCommand):
	def run(self,edit):
		self.view.run_command("expand_selection", {"to": "word"})
		sel = self.view.sel()[0]
		help_word = self.view.substr(sel)
		help_command = "help " + help_word

		StataAutomate(help_command)

class StataHelpInternal(sublime_plugin.TextCommand):
	def run(self,edit):
		self.view.run_command("expand_selection", {"to": "word"})
		sel = self.view.sel()[0]
		help_word = self.view.substr(sel)
		help_word = re.sub(" ","_",help_word)

		help_address = "http://www.stata.com/help.cgi?" + help_word
		helpfile_path = os.path.join(tempfile.gettempdir(), 'st_stata_help.txt')

		print(help_address)

		try:
			a = urllib.request.urlopen(help_address)
			source_code = a.read().decode("utf-8")
			a.close()

			regex_pattern = re.findall("<!-- END HEAD -->\n(.*?)<!-- BEGIN FOOT -->", source_code, re.DOTALL)
			help_content = re.sub("<h2>|</h2>|<pre>|</pre>|<p>|</p>|<b>|</b>|<a .*?>|</a>|<u>|</u>|<i>|</i>","",regex_pattern[0])
			help_content = re.sub("&gt;",">",help_content)
			help_content = re.sub("&lt;",">",help_content)

			with open(helpfile_path, 'w') as f:
				f.write(help_content)

			self.window = sublime.active_window()
			self.window.open_file(helpfile_path)
		
		except:
			print("Could not retrieve help file")

class StataLoad(sublime_plugin.TextCommand):
	def run(self,edit):
		sel = self.view.substr(self.view.sel()[0])
		StataAutomate("use " + sel + ", clear")

class StataLocal(sublime_plugin.TextCommand):
	def run(self,edit):
		sels = self.view.sel()
		for sel in sels:
			word_sel = self.view.word(sel.a)
			word_str = self.view.substr(word_sel)
			word_str = "`"+word_str+"'"
			self.view.replace(edit,word_sel,word_str)

class StataUpdateExecutablePath(sublime_plugin.ApplicationCommand):
	def run(self, **kwargs):

		def update_settings(fn):
			settings_fn = 'StataEditor.sublime-settings'
			settings = sublime.load_settings(settings_fn)

			old_fn = settings.get('stata_path', '')
			if old_fn!=fn:
				settings.set('stata_path_old', old_fn)

			settings.set('stata_path', fn)
			sublime.save_settings(settings_fn)
			sublime.status_message("Stata path updated")

		def cancel_update():
			sublime.status_message("Stata path not updated")

		def check_correct(fn):
			pass

		fn = get_exe_path()
		msg ="Enter the path of the Stata executable"
		sublime.active_window().show_input_panel(msg, fn, update_settings, check_correct, cancel_update)


# -------------------------------------------------------------
# Functions for Automation
# -------------------------------------------------------------

def getView():
	win = sublime.active_window()
	return win.active_view()
	
def get_cwd(view):
	fn = view.file_name()
	if not fn: return
	cwd = os.path.split(fn)[0]
	return cwd

def get_metadata(view):
	buf = sublime.Region(0, view.size())
	lines = (view.substr(line).strip() for line in view.split_by_newlines(buf))
	lines = [line[2:].strip() for line in lines if line.startswith('*!')]
	ans = {}
	for line in lines:
		key,val = line.split(':', 1)
		key = key.strip()
		# Allow dtapath instead of dtapaths
		if key=='dtapath': key = 'dtapaths'
		if key not in ans:
			ans[key] = [cell.strip() for cell in val.split(',')]
		# Allow repeated 'dtapaths' tags
		elif key=='dtapaths':
			ans[key].extend(cell.strip() for cell in val.split(','))
		else:
			print("Warning - Repeated autocomplete key:", key)
	if 'json' in ans: ans['json'] = ans['json'][0]
	ans['autoupdate'] = ans['autoupdate'][0].lower() in ('true','1','yes') if 'autoupdate' in ans else False
	if stata_debug: print('[METADATA]', ans)
	return ans

def get_autocomplete_data(view, force_update=False, add_from_buffer=True, obtain_varnames=True):

	# Will always check if there are new datasets in the given paths (except if autoupdate=False)
	# But will not update the varlists if all the datasets were modified before the last update

	# datasets is a tuple of (filename, pretty_dta_name)
	# variables is a tuple of (varname, pretty_var_name)

	is_stata = view.match_selector(0, "source.stata")
	if not is_stata:
		return
			
	cwd = get_cwd(view)
	if cwd is None:
		return
		
	metadata = get_metadata(view)
	paths = metadata.get('dtapaths', [])
	json_fn = metadata.get('json', '')
	json_fn = os.path.join(cwd, json_fn) if cwd and json_fn else ''
	json_exists = os.path.isfile(json_fn)
	autoupdate = True if force_update or not json_exists else metadata['autoupdate']

	if force_update and not json_fn:
		sublime.status_message('StataEdit Error: JSON filename was not set or file not saved')
		raise Exception(".json filename not specified in metadata")

	if json_exists and not force_update:
		# Read JSON
		with open(json_fn) as fh:
			data = json.load(fh)
		# If possible, use results stored in JSON
		if not autoupdate:
			variables = data['variables']
			datasets = data['datasets']

	# Else, first get list of datasets
	if autoupdate:
		datasets = get_datasets(view, paths)
	if stata_debug: print('[DATASETS]', datasets)

	# Get list of varnames
	if obtain_varnames and autoupdate:
		assert datasets # Bugbug
		if json_exists and not force_update:
			last_updated = data['updated']
			last_modified = max(os.path.getmtime(fn) for fn,_ in data['datasets'])
			needs_update = (last_updated<last_modified) or (datasets!=data['datasets'])
		else:
			needs_update = True
		variables = get_variables(datasets) if needs_update else data['variables']

	# Save JSON
	if autoupdate and json_fn and obtain_varnames and needs_update:
		last_updated = calendar.timegm(time.gmtime())
		data = {'updated': last_updated, 'datasets': datasets, 'variables': variables}
		with open(json_fn,'w') as fh:
			json.dump(data, fh, indent="\t")
		print('JSON file updated')

	# Add datasets from -save- commands and variables from -gen- commands
	if add_from_buffer:
		if obtain_varnames:
			if stata_debug: print('Varnames from current file:', get_generates(view))
			variables[' (current)'] = get_generates(view)
		else:
			datasets.extend(get_saves(view))

	if obtain_varnames:
		assert variables
	return (variables if obtain_varnames else datasets)

def get_datasets(view, paths):
	return list([fn,dta] for (fn,dta) in set( dta for path in paths for dta in get_dta_in_path(view, path) ))

def get_dta_in_path(view, path):
	"""Return list of tuples (full_filename, pretty_filename)"""

	# Paths may be relative to current Stata file
	cwd = get_cwd(view)
	os.chdir(cwd)

	nick = ''
	if '=' in path:
		nick, path = path.split('=', 1)
	if not os.path.isdir(path): return []
	# full file path, file name used in stata ($; no .dta)
	ans = [fn for fn in os.listdir(path) if fn.endswith('.dta')]
	ans = [ (os.path.join(path,fn), (nick if nick else path) + '/' + fn[:-4]) for fn in ans]
	return ans

def get_variables(datasets):
	"""Return dict of lists dta:varnames for all datasets"""
	return {dta:get_vars(fn) for (fn,dta) in datasets}

def get_vars(fn):
	# "use in 1" is too slow; just do "desc, varlist" 
	#cmd = """use "{}" in 1, clear nolabel""" 
	cmd = "describe using {}, varlist"
	StataAutomate(cmd.format(fn), sync=True)
	varlist = sublime.stata.StReturnString("r(varlist)")
	#sortlist = sublime.stata.StReturnString("r(sortlist)")
	#vars = sublime.stata.VariableNameArray()
	if stata_debug: print('[DTA={}]'.format(fn), varlist)
	return varlist.split(' ')

def get_saves(view):
	buf = sublime.Region(0, view.size())
	pat = '''^[ \t]*save[ \t]+"?([a-zA-Z0-9_`'.~ /:\\-]+)"?'''
	source = view.substr(buf)
	regex = re.findall(pat, source, re.MULTILINE)
	ans = [('',fn) for fn in regex]
	return ans

def get_generates(view):
	buf = sublime.Region(0, view.size())
	# Only accepts gen|generate|egen (the most common ones) and only the common numeric types
	pat = '''^[ \t]*(?:gen|generate|egen)[ \t]+(?:(?:byte|int|long|float|double)[ \t]+)?([a-zA-Z0-9_`']+)[ \t]*='''
	source = view.substr(buf)
	regex = re.findall(pat, source, re.MULTILINE)
	return list(set(regex))

# -------------------------------------------------------------
# Functions for Talking to Stata
# -------------------------------------------------------------

def plugin_loaded():
	global settings
	settings = sublime.load_settings(settings_file)

def StataAutomate(stata_command, sync=False):
	""" Launch Stata (if needed) and send commands """
	try:
		sublime.stata.DoCommand(stata_command) if sync else sublime.stata.DoCommandAsync(stata_command)
	except:
		launch_stata()
		sublime.stata.DoCommand(stata_command) if sync else sublime.stata.DoCommandAsync(stata_command)
	if stata_debug: print('[CMD]', stata_command)

def launch_stata():
	win32api.WinExec(settings.get("stata_path"), win32con.SW_SHOWMINNOACTIVE)
	sublime.stata = win32com.client.Dispatch ("stata.StataOLEApp")

	# Stata takes a while to start and will silently discard commands sent until it finishes starting
	# Workaround: call a trivial command and see if it was executed (-local- in this case)
	seed = int(random.random()*1e6) # Any number
	for i in range(50):
		sublime.stata.DoCommand('local {} ok'.format(seed))
		sublime.stata.DoCommand('macro list')
		rc = sublime.stata.MacroValue('_{}'.format(seed))
		if rc=='ok':
			sublime.stata.DoCommand('local {}'.format(seed)) # Empty it
			sublime.stata.DoCommand('cap cls')
			print("Stata process started (waited {}ms)".format((1+i)/10))
			break
		else:
			time.sleep(0.1)
	else:
		raise IOError('Stata process did not start before timeout')

def get_exe_path():
	reg = winreg.ConnectRegistry(None,winreg.HKEY_CLASSES_ROOT)
	try:
		key = winreg.OpenKey(reg, r"Applications\StataMP64.exe\shell\open\command")
		fn = winreg.QueryValue(key, None).strip('"').split('"')[0]
	except:
		print("Couldn't find path")
		return ''

	print(fn)
	return fn
