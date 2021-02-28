#!/usr/bin/env python
#--coding: utf-8 --
"""
Austin's Bulk Timeline Exporter
A tool to help export a large number of timelines quickly.
"""

import sys


# Stolen from python_get_resolve.py in the examples folder.
def GetResolve():
    try:
    # The PYTHONPATH needs to be set correctly for this import statement to work.
    # An alternative is to import the DaVinciResolveScript by specifying absolute path (see ExceptionHandler logic)
        import DaVinciResolveScript as bmd
    except ImportError:
        if sys.platform.startswith("darwin"):
            expectedPath="/Library/Application Support/Blackmagic Design/DaVinci Resolve/Developer/Scripting/Modules/"
        elif sys.platform.startswith("win") or sys.platform.startswith("cygwin"):
            import os
            expectedPath=os.getenv('PROGRAMDATA') + "\\Blackmagic Design\\DaVinci Resolve\\Support\\Developer\\Scripting\\Modules\\"
        elif sys.platform.startswith("linux"):
            expectedPath="/opt/resolve/libs/Fusion/Modules/"

        # check if the default path has it...
        print("Unable to find module DaVinciResolveScript from $PYTHONPATH - trying default locations")
        try:
            import imp
            bmd = imp.load_source('DaVinciResolveScript', expectedPath+"DaVinciResolveScript.py")
        except ImportError:
            # No fallbacks ... report error:
            print("Unable to find module DaVinciResolveScript - please ensure that the module DaVinciResolveScript is discoverable by python")
            print("For a default DaVinci Resolve installation, the module is expected to be located in: "+expectedPath)
            sys.exit()

    return bmd.scriptapp("Resolve")




resolve = GetResolve()
fusion = resolve.Fusion()
ui = fusion.UIManager



class AustinsBulkExporter:
	winID = 'com.austinwitherspoon.resolve.AustinsBulkExporter'
	window = None
	dispatcher = None
	project = None


	def __init__(self):
		global resolve
		self.buildUI()

		self.project = resolve.GetProjectManager().GetCurrentProject()

		self.populateSequences()
		self.populateRenderPresets()

		# Show window
		self.window.Show()
		self.dispatcher.RunLoop()


	# Build the interface
	def buildUI(self):
		global resolve, fusion, ui, bmd

		# If the window is already open, cancel this and bring it to the front.
		# Only one instance should exist at a time.
		self.window = ui.FindWindow(self.winID)
		if self.window:
			self.window.Show()
			self.window.Raise()
			exit()

		self.dispatcher = bmd.UIDispatcher(ui)

		# Build interface
		self.window = self.dispatcher.AddWindow({
			'ID': self.winID,
			'Geometry': [ 100,100,400,700 ],
			'WindowTitle': "Austin's Bulk Timeline Exporter",
			},
			ui.VGroup([
				ui.Label({ 'Text': "Bulk Timeline Exporter", 'Weight':0, 'Font': ui.Font({ 'Family': "Verdana", 'PixelSize': 20 }) }),
				ui.VGap(20, 0),
				ui.ComboBox({'ID': 'RenderPreset', 'Text':'Render Preset'}),
				ui.CheckBox({'ID': 'CutOffSlate', 'Text':'Force Timeline In to 01:00:00:00 (Cut off Slate)'}),
				ui.VGap(15, 0),
				ui.Label({'Text': 'Sequences', 'Weight':0}),
				ui.Tree({'ID': 'SequenceTree'}),
				ui.Button({'ID': 'Submit', 'Text': 'Add Selected Timelines to Queue!', 'Weight':0}),
				ui.Label({'Text': "Made by Austin ♥", 'Weight':0, 'Font':ui.Font({'PixelSize': 10})})
			])
		)


		# Register Events
		self.window.On[self.winID].Close = self.closeEvent
		self.window.On['SequenceTree'].ItemClicked = self.cleanTreeSelection
		self.window.On['Submit'].Clicked = self.submitRenders


	def submitRenders(self, event):
		preset = self.window.Find('RenderPreset').CurrentText
		cutOffSlate = self.window.Find('CutOffSlate').Checked
		
		if preset != 'Current Settings':
			self.project.LoadRenderPreset(preset)

		for timeline in self.selectedTimelines():
			self.project.SetCurrentTimeline(timeline)
			self.project.SetRenderSettings({'SelectAllFrames': True})

			if cutOffSlate:
				fps = int(round(float(timeline.GetSetting('timelineFrameRate'))))

				# Resolve's metadata appears to be wrong for non integer frame rates (except 23.98)
				if fps in [23, 29, 47, 59, 95, 119]:
					fps += 1

				frameIn = fps * 60 * 60

				self.project.SetRenderSettings({'MarkIn':frameIn})

			self.project.AddRenderJob()





	# Populate dropdown with render presets
	def populateRenderPresets(self):
		global resolve, ui
		dropdown = self.window.Find('RenderPreset')

		renderPresets = self.project.GetRenderPresetList()
		renderPresets.insert(0, 'Current Settings')

		dropdown.AddItems(renderPresets)


	# Scan for timelines in project and add to the Tree widget
	def populateSequences(self):
		global resolve, ui
		tree = self.window.Find('SequenceTree')

		# Make multiple items selectable
		tree.SetSelectionMode('ExtendedSelection')

		tree.SetHeaderHidden(True)

		# Make a root item
		sequencesItem = tree.NewItem()
		sequencesItem.Text[0] = 'Master'
		tree.AddTopLevelItem(sequencesItem)

		self.addItemsToTree(self.project.GetMediaPool().GetRootFolder(), sequencesItem)


	# Function to make it easier getting selected timelines from widget.
	# Will return a list of resolve Timeline objects
	def selectedTimelines(self):
		timelinesinProject = [self.project.GetTimelineByIndex(i+1) for i in range(self.project.GetTimelineCount())]
		selected = []
		tree = self.window.Find('SequenceTree')
		for item in tree.SelectedItems().values():
			if item.ChildCount() == 0:
				selected.append([i for i in timelinesinProject if i.GetName() == item.Text[0]][0])
		return selected


	# Recursive function to scan bins in resolve for timelines, and show all folders
	# containing timelines.
	def addItemsToTree(self, folder, rootItem):
		tree = rootItem.TreeWidget()

		# Sort subfolders by name and add them to the widget
		for subfolder in sorted(folder.GetSubFolderList(), key=lambda i: i.GetName()):
			row = tree.NewItem()
			row.Text[0] = '[' + subfolder.GetName() + ']'
			rootItem.AddChild(row)
			
			# If this folder ended up being empty, remove it
			if not self.addItemsToTree(subfolder, row):
				rootItem.RemoveChild(row)

		# Generate a list of all timelines in this folder (not timeline objects though - MPItems)
		mpItems = [i.GetName() for i in folder.GetClipList() if i.GetClipProperty('Type') == 'Timeline']
		# Get a list of all actual timeline items
		timelinesinProject = [self.project.GetTimelineByIndex(i+1) for i in range(self.project.GetTimelineCount())]
		# Match timelines to corresponding media pool items by name
		timelinesInFolder = [i for i in timelinesinProject if i.GetName() in mpItems]

		# Sort sequences and add them
		for sequence in sorted(timelinesInFolder, key=lambda i: i.GetName()):
			name = sequence.GetName()
			row = tree.NewItem()
			row.Text[0] = name
			rootItem.AddChild(row)

		rootItem.SetExpanded(True)

		# Return True if this folder is not empty (contains either folders or timelines)
		return rootItem.ChildCount() > 0

	def cleanTreeSelection(self, event):
		self.selectedTimelines()
		tree = self.window.Find('SequenceTree')
		for item in tree.SelectedItems().values():
			if item.ChildCount() > 0:
				for childID in range(item.ChildCount()):
					child = item.Child(childID)
					if child.ChildCount() == 0:
						child.SetSelected(True)



	def closeEvent(self, event):
		self.dispatcher.ExitLoop()


AustinsBulkExporter()