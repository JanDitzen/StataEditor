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

def StataAutomate(stata_command):
	""" Launch Stata (if needed) and send commands """
	try:
		sublime.stata.DoCommandAsync(stata_command)
	except:
		win32api.WinExec(settings.get("stata_path"))
		sublime.stata = win32com.client.Dispatch ("stata.StataOLEApp")
		sublime.stata.DoCommandAsync(stata_command)

class StataExecuteCommand(sublime_plugin.TextCommand):
	def get_path(self):
		fn = self.window.active_view().file_name()
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
		
		cwd = get_path()
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
