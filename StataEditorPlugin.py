import sublime, sublime_plugin
import os
import Pywin32.setup
import win32com.client
import win32con
# http://msdn.microsoft.com/en-us/library/windows/desktop/ms633548(v=vs.85).aspx
import win32api
import tempfile
import subprocess
import re
import urllib
from urllib import request
import hashlib
import json
import random
import time, calendar

settings_file = "StataEditor.sublime-settings"

def plugin_loaded():
	global settings
	settings = sublime.load_settings(settings_file)

# def StataRunning():
# 	""" Check if Stata is running """
# 	cmd = "WMIC PROCESS get Caption"
# 	proc = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)

# 	all_run_prog = ""
# 	for line in proc.stdout:
# 		all_run_prog = all_run_prog + str(line) + "\n"

# 	prog_run = re.findall('Stata.*?\.exe', all_run_prog)

# 	if len(prog_run) > 0:
# 		return True
# 	else:
# 		return False

# http://www.stata.com/automation/
# To get locals and globals:
# sublime.stata.MacroValue(aGlobal)
# sublime.stata.MacroValue(_aLocal)
# Also.. scalar=ScalarType
# StReturnString("c(current_date)") StReturnType("c(current_date)") StReturnNumeric("c(max_matsize)")
# UtilGetStMissingValue
# StVariableName(#) #>=1 <=c(K)
# VariableType VariableNameArray !

def StataAutomate(stata_command, sync=False):
	""" Launch Stata (if needed) and send commands """
	# method = sublime.stata.DoCommand if sync else sublime.stata.DoCommandAsync
	try:
		sublime.stata.DoCommand(stata_command) if sync else sublime.stata.DoCommandAsync(stata_command)
	except:
		launch_stata()
		sublime.stata.DoCommand(stata_command) if sync else sublime.stata.DoCommandAsync(stata_command)

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

# Cambiar el completion por un quick panel show_quick_panel 
# https://gist.github.com/robmccormack/6040840
# http://stackoverflow.com/questions/12976008/accessing-the-quick-panel-in-a-sublime-text-2-plugin
# En el caso de variables, escoger el dta, y luego escoger variables anhadiendo espacio hasta que aprete escape..
# Hacer override de: { "keys": ["ctrl+shift+space"], "command": "expand_selection", "args": {"to": "scope"} },
# self.window.show_quick_panel([[cmd["title"], cmd["command"]] for cmd in self._commands] + [["<New>", "Create a new command"]],
#self._on_select_command_done)

def get_cwd(view):
	fn = view.file_name()
	if not fn: return
	cwd = os.path.split(fn)[0]
	return cwd

def read_json(fn):
	if not fn or not os.path.isfile(fn): return None
	with open(fn) as fh:
		d = json.load(fh)
		return d

def get_metadata(view):
	buf = sublime.Region(0, view.size())
	lines = [view.substr(line).strip() for line in view.split_by_newlines(buf)]
	lines = [line[2:].strip() for line in lines if line.startswith('*!')]
	ans = {}
	for line in lines:
		key,val = line.split(':', 1)
		ans[key.strip()] = [cell.strip() for cell in val.split(',')]
	return ans

def get_saves(view):
	buf = sublime.Region(0, view.size())
	pat = '''^[ \t]*save[ \t]+"?([a-zA-Z0-9_`'.~ /:\\-]+)"?'''
	source = view.substr(buf)
	regex = re.findall(pat, source, re.MULTILINE)
	ans = [('',fn) for fn in regex]
	return ans

def get_dta_in_path(path):
	nick = ''
	if '=' in path:
		nick, path = path.split('=', 1)
	if not os.path.isdir(path): return []
	# full file path, file name used in stata ($; no .dta)
	ans = [fn for fn in os.listdir(path) if fn.endswith('.dta')]
	ans = [ (os.path.join(path,fn), (nick if nick else path) + '/' + fn[:-4]) for fn in ans]
	return ans

def prepare_dta_suggestion(dta):
	#desc = 'temp' if r"`" in dta[1] else 'dta'
	return dta[1]
	#return '[{}]: {}'.format(desc, dta[1])

def get_dta_suggestions(view):
	metadata = get_metadata(view)
	paths = metadata.get('dtapaths', [])
	datasets = get_saves(view)
	for path in paths:
		datasets.extend(get_dta_in_path(path))
	datasets = tuple(set(datasets))
	return datasets, [prepare_dta_suggestion(dta) for dta in datasets]

def get_var_suggestions(view):
	metadata = get_metadata(view)
	paths = metadata.get('dtapaths', [])
	# Don't use get_dta_suggestions b/c I don't want suggestions from -save-
	datasets = []
	for path in paths:
		datasets.extend(get_dta_in_path(path))
	datasets = tuple(set(datasets))
	dta_files = tuple(dta[0] for dta in datasets)
	dta_stata = tuple(dta[1] for dta in datasets)
	hashvalue = hash(dta_files)

	cwd = get_cwd(view)
	json_fn = metadata.get('json', [''])[0]
	json_fn = os.path.join(cwd, json_fn) if cwd and json_fn else ''

	if json_fn:
		data = read_json(json_fn) # Read old json
		if data is None: data = {'modified': 0}
		#old_hashvalue = data['hashvalue'] if data else None # If old json is stale, refresh
		#if old_hashvalue!=hashvalue:
		if max(os.path.getmtime(fn) for fn in dta_files) > data['modified']:
			data = save_json(json_fn, datasets, hashvalue) # Save and update data dict
			print('JSON file updated')
		return zip(*data['vars']) # obtain dta, suggestion pairs and transform into 2 lists
	else:
		return [],[]

def save_json(json_fn, datasets, hashvalue):
	modified = calendar.timegm(time.gmtime())
	data = {'hashvalue':hashvalue, 'modified':modified}
	variables = []
	for fn, dta in datasets:
		varnames = get_vars(fn)
		variables.extend( (var, '[{}]: {}'.format(dta, var)) for var in varnames)
	data['vars'] = variables
	with open(json_fn,'w') as fh:
		json.dump(data, fh, indent="\t")
	return data

def get_vars(fn):
	# "use in 1" is too slow; just do "desc, varlist" 
	#cmd = """use "{}" in 1, clear nolabel""" 
	cmd = "describe using {}, varlist"
	print(cmd.format(fn))
	StataAutomate(cmd.format(fn), sync=True)
	varlist = sublime.stata.StReturnString("r(varlist)")
	#sortlist = sublime.stata.StReturnString("r(sortlist)")
	#vars = sublime.stata.VariableNameArray()
	return varlist.split(' ')

class StataAutocompleteDtaCommand(sublime_plugin.TextCommand):
	def run(self, edit):
		self.datasets, self.suggestions = get_dta_suggestions(self.view)
		self.view.window().show_quick_panel(self.suggestions, self.insert_link) #, sublime.MONOSPACE_FONT)

	def insert_link(self, choice):
		if choice==-1:
			return
		link = '"' + self.datasets[choice][1] + '"'
		self.view.run_command("stata_insert", {'link':link})

class StataAutocompleteVarCommand(sublime_plugin.TextCommand):
	def run(self, edit):
		self.varlist, self.suggestions = get_var_suggestions(self.view)
		if self.varlist:
			self.view.window().show_quick_panel(self.suggestions, self.insert_link) #, sublime.MONOSPACE_FONT)

	def insert_link(self, choice):
		if choice==-1:
			return
		link = self.varlist[choice] + ' '
		self.view.run_command("stata_insert", {'link':link})
		# Call again except if ESC is pressed
		sublime.set_timeout(lambda: self.view.run_command("stata_autocomplete_var"), 1)

class StataInsert(sublime_plugin.TextCommand):
	def run(self, edit, link):
		startloc = self.view.sel()[-1].end()
		self.view.insert(edit, startloc, link)

class StataExecuteCommand(sublime_plugin.TextCommand):
	def get_path(self):
		fn = self.view.window().active_view().file_name()
		return None if not fn else os.path.split(fn)[0]

	def run(self, edit, **args):
		all_text = ""
		len_sels = 0
		sels = self.view.sel()
		len_sels = 0
		for sel in sels:
			len_sels = len_sels + len(sel)

		if len_sels == 0:
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
		
		cwd = self.get_path()
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

		help_adress = "http://www.stata.com/help.cgi?" + help_word
		helpfile_path = os.path.join(tempfile.gettempdir(), 'st_stata_help.txt')

		print(help_adress)

		try:
			a = urllib.request.urlopen(help_adress)
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
