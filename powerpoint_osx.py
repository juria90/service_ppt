'''powerpoint_osx.py implements App and Presentation for Powerpoint in OS X using applescript.

Powerpoint AppleScript reference: http://www.codemunki.com/PPT2004AppleScriptRef.pdf
https://stackoverflow.com/questions/7570855/where-to-find-what-commands-properties-are-available-for-applescript-in-microsof

Convert VBA to AppleScript: http://preserve.mactech.com/vba-transition-guide/index-toc.html
https://developer.microsoft.com/en-us/microsoft-365/blogs/vba-improvements-in-office-2016/

AppleScript to select eraser in Powerpoint: https://stackoverflow.com/questions/60341314/use-applescript-to-select-eraser-in-powerpoint

AppleScript: https://nathangrigg.com/images/2012/AppleScriptLanguageGuide.pdf


AST: https://apple.stackexchange.com/questions/249606/how-to-code-an-applescript-to-do-a-mouse-click-on-a-menu-bar-item
cliclick: https://apple.stackexchange.com/questions/323916/applescript-click-by-position
'''

import os
import subprocess
import traceback

# pip install py-applescript
import applescript

from powerpoint_base import SlideCache, PresentationBase, PPTAppBase


AEKEY_KEYDATA = applescript.AEType(b'seld')


def get_KeyData(ae):
    seld_k = AEKEY_KEYDATA
    seld = ae[seld_k]
    return seld

def run_applescript(prefix_cmd, cmd, reraise_exception=False):
    try:
        full_cmd = ''
        if prefix_cmd:
            full_cmd = prefix_cmd + '\n'
        full_cmd = full_cmd + 'tell app "Microsoft PowerPoint"\n' + cmd + '\nend tell'

        scpt = applescript.AppleScript(full_cmd)
        result = scpt.run()

        return result
    except applescript.ScriptError as e:
        print("Error: %s" % str(e))
        traceback.print_exc()

        if reraise_exception:
            raise


class Presentation(PresentationBase):
    '''Presentation is a wrapper around PowerPoint.Presentation object.
    All the index used in the function is 0-based and converted to 1-based while calling COM functions.
    '''

    def __init__(self, app, prs):
        super().__init__(app, prs)

    def slide_count(self):
        prs_name = get_KeyData(self.prs)
        cmd = f'set slide_count to count of slides of presentation "{prs_name}"\nreturn slide_count'
        slide_count = run_applescript('', cmd)
        return slide_count

    def _fetch_slide_cache(self, slide_index):
        sc = SlideCache()

        sc.id = self.slide_index_to_ID(slide_index)

        sc.slide_text = self._get_texts_in_shapes(slide_index)
        sc.notes_text = self._get_texts_in_note_shapes(slide_index)

        sc.valid_cache = True

        return sc

    def slide_index_to_ID(self, var):
        prs_name = get_KeyData(self.prs)
        if isinstance(var, int):
            sindex = var + 1
            cmd = f'set slide_id to slide ID of slide {sindex} of presentation "{prs_name}"\nreturn slide_id'
            slide_id = run_applescript('', cmd)
            return slide_id
        elif isinstance(var, list):
            result = []
            for index in var:
                sindex = index + 1
                cmd = f'set slide_id to slide ID of slide {sindex} of presentation "{prs_name}"\nreturn slide_id'
                slide_id = run_applescript('', cmd)
                result.append(slide_id)

            return result

    def slide_ID_to_index(self, var):
        '''slide_ID_to_index() does not work because the syntax to get the slide ID is not accurate, yet.
        '''
        prs_name = get_KeyData(self.prs)
        if isinstance(var, int):
            cmd = f'set slide_index to item 1 of (get every slide of presentation "{prs_name}" whose slide ID is {var})\nreturn slide_index'
            result = run_applescript('', cmd)
            slide_index = get_KeyData(result)
            return slide_index - 1
        elif isinstance(var, list):
            slide_indices = []
            for sid in var:
                cmd = f'set slide_index to item 1 of (get every slide of presentation "{prs_name}" whose slide ID is {var})\nreturn slide_index'
                result = run_applescript('', cmd)
                slide_index = get_KeyData(result)
                slide_indices.append(slide_index - 1)

            return result

    def _get_texts_in_shapes(self, slide_index):
        prs_name = get_KeyData(self.prs)
        sindex = slide_index + 1
        prefix_cmd = 'set theList to {}'
        cmd = f'''
tell slide {sindex} of presentation "{prs_name}"
  repeat with shapeNumber from 1 to count of shapes
    set shapeText to content of text range of text frame of shape shapeNumber
    set end of theList to shapeText
  end repeat
  return theList
end tell
'''
        theList = run_applescript(prefix_cmd, cmd)
        theList = [s for s in theList if isinstance(s, str)]
        return theList

    def _get_texts_in_note_shapes(self, slide_index):
        prs_name = get_KeyData(self.prs)
        sindex = slide_index + 1
        prefix_cmd = 'set theList to {}'
        cmd = f'''
tell notes page of slide {sindex} of presentation "{prs_name}"
  repeat with shapeNumber from 1 to count of shapes
    set shapeText to content of text range of text frame of shape shapeNumber
    set end of theList to shapeText
  end repeat
  return theList
end tell
'''
        theList = run_applescript(prefix_cmd, cmd)
        theList = [s for s in theList if isinstance(s, str)]
        return theList

    def _replace_texts_in_slide_shapes(self, slide_index, find_replace_dict):
        '''_replace_texts_in_slide_shapes() utilize "System Events" extensively to implement Find and Replace.
        Please refer applescript/ppt-findReplace_slide.scpt for more information.
        '''
        self.activate()

        # create AppleScript list from find_replace_dict
        find_texts = '{' + ','.join(['"' + key + '"' for key in find_replace_dict.keys()])+ '}'
        replace_texts = '{' + ','.join(['"' + value + '"' for _k, value in find_replace_dict.items()])+ '}'
        prs_name = get_KeyData(self.prs)
        slide_index_1 = slide_index + 1

        cmd = f'''
set find_texts to {find_texts}
set replace_texts to {replace_texts}
set prs_name to "{prs_name}"
set slide_index to {slide_index_1}

set findnext_label to "Find Next"
set replace_label to "Replace"
set done_label to "Done"

on get_current_slide_index(prs_name)
	set slide_index to -1
	tell application "Microsoft PowerPoint"
		tell presentation prs_name
			set slide_index to slide index of slide range of selection of document window 1
		end tell
	end tell

	return slide_index
end get_current_slide_index

on set_selected_slide(prs_name, slide_index)
	tell application "Microsoft PowerPoint"
		log "Selecting slide " & slide_index
		set theView to view of document window 1
		go to slide theView number slide_index
	end tell
end set_selected_slide

on cannot_find_any_more_in_slide_index(prs_name, slide_index)
	tell application "System Events"
		tell process "Microsoft PowerPoint"
			# wait for "We couldn't find what you were looking for." dialog box.
			# button OK of window 1
			repeat 3 times
				# log "repeat"
				if (name of window 1 = "") then
					if (exists button "OK" of window 1) then
						click button "OK" of window 1
						delay 1
						return true
					end if
				end if

				# if new selected slide is not same as original, exit.
				set cur_slide_index to my get_current_slide_index(prs_name)
				if cur_slide_index is not slide_index then return true

				delay 0.5
			end repeat
		end tell
	end tell

	return false
end cannot_find_any_more_in_slide_index

on replace_text_in_slide(prs_name, slide_index, find_text, replace_text, replace_label, done_label)
	tell application "System Events"
		tell process "Microsoft PowerPoint"
			# Replace dialog cuts replace_text when it is longer than 128.
			# So, close dialog and paste manually, then bring up the dialog again.

			set strLength to the length of replace_text
			log "replace with paste"
			click button done_label of window "Replace"
			delay 1

			# Paste and Match Formatting.
			keystroke "v" using {{command down, option down, shift down}}
			delay 0.5

			keystroke "h" using {{control down}}
			delay 1
		end tell
	end tell
end replace_text_in_slide

tell application "System Events"
	tell process "Microsoft PowerPoint"
		set frontmost to true
		delay 1

		# click View -> Slide Sorter menu
		keystroke "1" using {{command down}}
		delay 1

		# click Edit -> Find -> Replace... menu
		keystroke "h" using {{control down}}
		delay 1

		set item_index to 0
		repeat with find_text in find_texts
			# item_index start from 1 in applescript
			set item_index to item_index + 1

			set replace_text to item item_index of replace_texts

			my set_selected_slide(prs_name, slide_index)

			# Find What:
			set the clipboard to {{Unicode text:find_text}}

			# select combo box 1 of window "Replace"
			# Use copy/paste to handle any chars.
			keystroke "av" using {{command down}}
			delay 0.5

			# Replace With:
			set the clipboard to {{Unicode text:replace_text}}

			# select combo box 2 of window "Replace"
			# tell combo box 2 doesn't work, so use tab key to navigate to next combo box.
			keystroke (ASCII character 9)
			keystroke "av" using {{command down}}
			delay 0.5

			# move to find what: combo box
			keystroke (ASCII character 9)

			repeat
				click button findnext_label of window "Replace"
				delay 1

				set no_more to my cannot_find_any_more_in_slide_index(prs_name, slide_index)
				if no_more then
					exit repeat
				end if

				my replace_text_in_slide(prs_name, slide_index, find_text, replace_text, replace_label, done_label)

				# switch slide to previous or next and back to current to unselect shapes
				set cur_slide_index to my get_current_slide_index(prs_name)
				if cur_slide_index = slide_index then
					if slide_index is greater than 1 then
						my set_selected_slide(prs_name, slide_index - 1)
					else
						my set_selected_slide(prs_name, slide_index + 1)
					end if
				end if

				my set_selected_slide(prs_name, slide_index)
			end repeat

		end repeat

		click button done_label of window "Replace"
		delay 1
	end tell
end tell
'''
        try:
            scpt = applescript.AppleScript(cmd)
            scpt.run()
        except applescript.ScriptError:
            raise

    def _replace_all_slides_texts(self, find_replace_dict):
        '''_replace_all_slides_texts() utilize "System Events" extensively to implement Find and Replace.
        Please refer applescript/ppt-findReplaceAll.scpt for more information.
        '''
        self.activate()

        # create AppleScript list from find_replace_dict
        find_texts = '{' + ','.join(['"' + key + '"' for key in find_replace_dict.keys()])+ '}'
        replace_texts = '{' + ','.join(['"' + value + '"' for _k, value in find_replace_dict.items()])+ '}'

        cmd = f'''
set find_texts to {find_texts}
set replace_texts to {replace_texts}
set replaceall_label to "Replace All"
set done_label to "Done"

tell application "System Events"
	tell process "Microsoft PowerPoint"
		set frontmost to true

		# click Edit -> Find -> Replace... menu
		click menu item "Replace..." of menu 1 of menu item "Find" of menu "Edit" of menu bar 1
		delay 1

		set item_index to 0
		repeat with find_text in find_texts
			# item_index start from 1 in applescript
			set item_index to item_index + 1

			set replace_text to item item_index of replace_texts

			# Find What:
			set the clipboard to {{Unicode text:find_text}}

			# select combo box 1 of window "Replace"
			# Use copy/paste to handle any chars.
			keystroke "av" using {{command down}}
			delay 0.5

			# Replace With:
			set the clipboard to {{Unicode text:replace_text}}

			# select combo box 2 of window "Replace"
			# tell combo box 2 doesn't work, so use tab key to navigate to next combo box.
			keystroke (ASCII character 9)
			keystroke "av" using {{command down}}
			delay 0.5

			# move to find what: combo box
			keystroke (ASCII character 9)

			click button replaceall_label of window "Replace"
			delay 1

			# wait for "Powerpoint searched your presentation and made x replacements." dialog box
			# Or "We couldn't find what you were looking for." dialog box.
			# button OK of window 1
			repeat 5 times
				if (name of window 1 = "") then
					if (exists button "OK" of window 1) then
						click button "OK" of window 1
						exit repeat
					end if
				end if
				delay 1
			end repeat
		end repeat

		click button done_label of window "Replace"
		delay 1
	end tell
end tell
'''
        try:
            scpt = applescript.AppleScript(cmd)
            scpt.run()
        except applescript.ScriptError:
            raise

    def duplicate_slides(self, source_location, insert_location=None, copy=1):
        '''duplicate_slides() requires System Events that needs security set up.
        If the implementation can be done without it, it will be much robust.
        '''

        source_count = 1
        if isinstance(source_location, int):
            source_location = source_location + 1
            source_count = 1
        elif isinstance(source_location, list):
            if len(source_location) == 1:
                source_location = source_location[0] + 1
                source_count = 1
            else:
                # only support 1 slide duplication.
                raise ValueError('Invalid insert location')

                # make sure all elements are integer
                # source_location = [int(i) + 1 for i in source_location]
                # source_count = len(source_location)
        else:
            return 0

        # insert_location can change by copying slides.
        # So save the ID and use it later.
        if insert_location is None:
            insert_location = self.slide_count() - 1
        elif isinstance(insert_location, int):
            insert_location = insert_location + 1
        else:
            raise ValueError('Invalid insert location')

        added_count = 0

        prs_name = get_KeyData(self.prs)
        for _ in range(copy):
            source_location_1 = source_location + 1
            insert_location1 = insert_location
            if insert_location > source_location:
                insert_location1 = insert_location1 + source_count

            cmd = f'''
tell presentation "{prs_name}"
  activate
  select slide {source_location}
  tell application "System Events" to keystroke "D" using command down
  delay 0.1
  move slide {source_location_1} to before slide {insert_location1}
  delay 0.1
end tell
'''
            run_applescript('', cmd)

            added_count = added_count + source_count

            self._slides_inserted(insert_location - 1, source_count)
            if source_location >= insert_location:
                source_location = source_location + source_count

            insert_location = insert_location + source_count

        return added_count

    def insert_blank_slide(self, slide_index):
        insert_location = 'the beginning'
        if slide_index == 0:
            pass
        elif slide_index == self.slide_count():
            insert_location = 'the end'
        else:
            insert_location = f'after slide {slide_index}'
        prs_name = get_KeyData(self.prs)
        cmd = f'make new slide at {insert_location} of presentation "{prs_name}" with properties {{layout: slide layout blank}}'
        run_applescript('', cmd)

    def delete_slide(self, slide_index):
        prs_name = get_KeyData(self.prs)
        slide_index = slide_index + 1
        cmd = f'''
tell presentation "{prs_name}"
  delete slide {slide_index}
end tell
'''
        run_applescript('', cmd)

    def activate(self):
        # activate the presentation before using "System Events"
        prs_name = get_KeyData(self.prs)
        cmd = f'''
tell presentation "{prs_name}"
  activate
end tell
'''
        try:
            run_applescript('', cmd, reraise_exception=True)
        except applescript.ScriptError:
            raise

    def copyall_and_close(self):
        self.activate()

        cmd = f'''
tell application "System Events"
	tell process "Microsoft PowerPoint"
		set frontmost to true

		# click View -> Slide Sorter menu
		click menu item "Slide Sorter" of menu "View" of menu bar 1
		delay 1

		# Command + A to select all slides.
		# Command + C to copy slides.
		keystroke "ac" using {{command down}}
		delay 1

		# Command + W to close the window.
		keystroke "w" using {{command down}}
		delay 1
	end tell
end tell
'''
        try:
            scpt = applescript.AppleScript(cmd)
            scpt.run()
        except applescript.ScriptError:
            raise

    def _paste_keep_source_formatting(self, insert_location):
        '''_paste_keep_source_formatting() uses System Events to paste the copied slides in the clipboard.
        The pastes slides will be inserted after the selection set by "select slide insert_location".
        So, the logic needs to take care of that.

        For clicking "Paste with Keep Source Formatting", it doesn't work with AppleScript,
        because it's using AXPopupMenu.

        So, cliclick utility is used below, but can be replaced by other Python module like pyautogui.
        '''

        insert_location_1 = insert_location - 1
        prs_name = get_KeyData(self.prs)
        cmd = \
f'''
set insert_location to {insert_location_1}

tell application "Microsoft PowerPoint"
	tell presentation "{prs_name}"
		activate
		select slide insert_location
	end tell
end tell

set keep_source_formatting to "Keep Source Formatting"

tell application "System Events"
	set arrow_x to 0
	set arrow_y to 0

	tell process "Microsoft PowerPoint"
		set frontmost to true

		set btn_pos to {0, 0}
		set btn_sz to {0, 0}
		tell menu button "Paste" of group 1 of scroll area 1 of tab group 1 of window 1
			set btn_pos to position
			set btn_sz to size

			set arrow_x to (item 1 of btn_pos) + (item 1 of btn_sz) - 6
			set arrow_y to (item 2 of btn_pos) + 10
		end tell

		# Use cliclick utility to generate physical click at the position.
		set cmd to "/usr/local/bin/cliclick c:" & (arrow_x as string) & "," & (arrow_y as string)
		do shell script cmd

		delay 1

		tell group 1 of window "Paste"
			click menu button keep_source_formatting
			delay 1
		end tell
	end tell
end tell
'''
        try:
            scpt = applescript.AppleScript(cmd)
            scpt.run()
        except applescript.ScriptError:
            raise

    def saveas(self, filename):
        try:
            prs_name = get_KeyData(self.prs)
            prefix_cmd = f'set f to POSIX file "{filename}"'
            cmd = f'save presentation "{prs_name}" in f as presentation'
            run_applescript(prefix_cmd, cmd, reraise_exception=True)
        except applescript.ScriptError:
            pass
        else:
            _dir, basename = os.path.split(filename)
            fn, _ext = os.path.splitext(basename)
            self.prs[AEKEY_KEYDATA] = fn

    def saveas_format(self, filename, image_type='png'):
        '''saveas_format() utilize "System Events" extensively to implement Export As JPEG/PNG.
        Please refer applescript/ppt-export.scpt for more information.
        '''
        self.activate()

        output_dir, leaf_dir = os.path.split(filename)
        file_format = 'PNG'
        if image_type == 'jpg':
            file_format = 'JPEG'

        cmd = \
f'''
set output_dir to "{output_dir}"
set leaf_dir to "{leaf_dir}"
set file_format to "{file_format}" # "PNG"
set radio_every to 1 # "Save Every Slide"
set radio_current to 2 # "Save Current Slide Only"
set button_export to "Export"
set button_replace to "Replace"

tell application "System Events"
	tell process "Microsoft PowerPoint"
		set frontmost to true
		delay 0.5

		# click File -> Export... menu
		click menu item "Export..." of menu "File" of menu bar 1
		delay 1

		# Command + Shift + G to bring up "Go to the folder:" dialog.
		keystroke "g" using {{command down, shift down}}
		delay 1

		keystroke output_dir
		delay 0.5

		click button "Go" of sheet 1 of sheet 1 of window 1
		delay 1

		# type the leaf directory
		tell text field "Export As:" of sheet 1 of window 1
			keystroke leaf_dir
			delay 0.5
		end tell

		# "File Format:" button
		tell pop up button 2 of sheet 1 of window 1
			click
			delay 1

			tell menu 1
				click menu item file_format
				delay 1
			end tell
		end tell

		# JPEG & PNG option.
		click radio button radio_every of radio group 1 of sheet 1 of window 1
		# click radio button radio_current of radio group 1 of sheet 1 of window 1
		delay 0.5

		# click Export button
		click button button_export of sheet 1 of window 1
		delay 1

		# click "Replace" button in “filename” already exists. Do you want to replace it?
		set replace_button to a reference to (first button whose name is button_replace) of sheet 1 of sheet 1 of window 1
		if replace_button exists then
			click replace_button
			delay 1
		end if

		# wait for "Each slide in your presentation has been saved as a separate file in the folder XYZ." dialog box
		# for 1 minute to finish.
		repeat 60 times
			# button OK of window 1
			if (exists button "OK" of window 1) then
				click button "OK" of window 1
				delay 1
				exit repeat
			end if

			delay 1
		end repeat
	end tell
end tell
'''
        try:
            scpt = applescript.AppleScript(cmd)
            scpt.run()
        except applescript.ScriptError:
            raise

    def close(self):
        prs_name = get_KeyData(self.prs)
        cmd = f'close presentation "{prs_name}"'
        run_applescript('', cmd)
        self.prs = None


class App(PPTAppBase):

    @staticmethod
    def is_running():
        process = subprocess.Popen('pgrep "Microsoft PowerPoint"', shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        my_pid, _err = process.communicate()

        try:
            pid = int(my_pid)
        except:
            return False

        return True

    def __init__(self):
        cmd = 'activate'
        run_applescript('', cmd)

    def new_presentation(self):
        cmd = 'set prs to make new presentation\nreturn prs'
        prs = run_applescript('', cmd)
        if prs:
            return Presentation(self, prs)

    def open_presentation(self, filename):
        _dir, basename = os.path.split(filename)
        fn, _ext = os.path.splitext(basename)

        prefix_cmd = f'set f to POSIX file "{filename}"'
        cmd = \
f'''open f
set prs to presentation "{fn}"
return prs'''
        prs = run_applescript(prefix_cmd, cmd)
        if prs:
            return Presentation(self, prs)

    def presentation_count(self):
        cmd = 'set pcount to count of presentations\nreturn pcount'
        pcount = run_applescript('', cmd)
        return pcount

    def quit_if_empty(self):
        presentation_count = self.presentation_count()
        if presentation_count == 0:
            self.quit()

    def quit(self):
        cmd = 'quit'
        run_applescript('', cmd)
