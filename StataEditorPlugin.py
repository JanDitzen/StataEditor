import sublime, sublime_plugin
import os
import Pywin32.setup
import win32com.client
import win32api
import tempfile
import subprocess
import re
import urllib
from urllib import request
import hashlib
import json

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
		win32api.WinExec(settings.get("stata_path"))
		sublime.stata = win32com.client.Dispatch ("stata.StataOLEApp")
		sublime.stata.DoCommand(stata_command) if sync else sublime.stata.DoCommandAsync(stata_command)
	version = StReturnNumeric("c(stata_version)")
	print('Stata version:', version)

class StataDtaAutocompleteCommand(sublime_plugin.TextCommand):
	def run(self, edit, **args):
		metadata = self.get_metadata()
		paths = metadata.get('dtapaths', [])
		datasets = self.get_saves()
		for path in paths:
			datasets.extend(self.get_dta_in_path(path))
		datasets = tuple(set(datasets))
		hashvalue = hash(datasets)

		cwd = self.get_cwd()
		json_fn = metadata.get('json', [''])[0]
		json_fn = os.path.join(cwd, json_fn) if cwd and json_fn else ''

		# If there is a json name specified,
		# And the list of dtas is different, update the variables
		# BUGBUG: What if the names don't change but the contents do? In that case, delete the .json

		if json_fn:
			data = self.read_json(json_fn) # Read old json
			old_hashvalue = data['hashvalue'] if data else None # If old json is stale, refresh
			if old_hashvalue!=hashvalue or True: # Bugbug
				#sublime.error_message(str([old_hashvalue,hashvalue]))
				print('JSON file updated', old_hashvalue,hashvalue)
				data = self.save_json(json_fn, datasets, hashvalue) # Save and update data dict
		#sublime.error_message(json_fn)
		#sublime.error_message(str(datasets))
		# What do i do with -data- now??
		print('Done!')

	def get_cwd(self):
		fn = self.view.window().active_view().file_name()
		if not fn: return
		cwd = os.path.split(fn)[0]
		return cwd

	def read_json(self, fn):
		if not fn or not os.path.isfile(fn): return []
		with open(fn) as fh:
			d = json.load(fh)
			return d

	def save_json(self, fn, datasets, hashvalue):
		data = {'datasets':datasets, 'hashvalue':hashvalue}
		vars = {}
		for dta in datasets:
			if dta[0]:
				vars[dta[2]] = self.get_vars(os.path.join(dta[0],dta[1]))
		data['vars'] = vars
		with open(fn,'w') as fh:
			json.dump(data, fh, indent="\t")
		return data

	def get_vars(self, fn):
		cmd = """use "{}" in 1 if 0, clear nolabel"""
		StataAutomate(cmd.format(fn), sync=True)
		vars = sublime.stata.VariableNameArray()
		print(fn, vars)
		return vars

	def get_saves(self):
		buf = sublime.Region(0, self.view.size())
		pat = '''^[ \t]*save[ \t]+"?([a-zA-Z0-9_`'.~ /:\\-]+)"?'''
		source = self.view.substr(buf)
		regex = re.findall(pat, source, re.MULTILINE)
		ans = [('',fn,fn) for fn in regex]
		return ans

	def get_metadata(self):
		buf = sublime.Region(0, self.view.size())
		lines = [self.view.substr(line).strip() for line in self.view.split_by_newlines(buf)]
		lines = [line[2:].strip() for line in lines if line.startswith('*!')]
		ans = {}
		for line in lines:
			key,val = line.split(':', 1)
			ans[key.strip()] = [cell.strip() for cell in val.split(',')]
		return ans

	def get_dta_in_path(self, path):
		nick = ''
		if '=' in path:
			nick, path = path.split('=', 1)
		if not os.path.isdir(path): return []
		ans = [(path,fn, (nick if nick else path) + '/' + fn) for fn in os.listdir(path) if fn.endswith('.dta')]
		return ans

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
