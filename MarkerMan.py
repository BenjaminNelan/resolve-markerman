#!/user/bin/env python3
# Python 3.10.11

"""

DaVinci Resolve Marker Manager
Benjamin Nelan 2023

- Allows markers to be used to mark 'In / Out Points' of clips.

Python 3.10.11
DaVinci Resolve Studio 18.1.4

v0.1.1

# Forum discussion about scripts in the free version:
# https://forum.blackmagicdesign.com/viewtopic.php?f=21&t=113252
"""

import os, tempfile, errno, re, unicodedata, sys
import tkinter as tk

class MarkerManager:
    def __init__(self):
        self.bmd = GetBMD()
        self.resolve = app.GetResolve()
        self.fusion = app
        self.project = self.resolve.GetProjectManager().GetCurrentProject()
        self.timeline = self.project.GetCurrentTimeline()
        self.clips = []
        self.markers = {}
        self.screen = self.GetScreenSize()
        self.version = 0.11
        self.markerProcessingFunction = False
        self.renderLocation = False

        # Displays UI Prompt in Resolve
        # Seems to work externally even though posts say it's not meant to.
        try:
            self.ui = self.fusion.UIManager
            self.disp = self.bmd.UIDispatcher(self.ui)
            self.DialogSelectMarkerColor()
        except:
            self.ui = False
            self.disp = False

    def GetScreenSize(self):
        root = tk.Tk()
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        scaling_factor = root.tk.call('tk', 'scaling')
        root.quit()
        return {
            'width': screen_width,
            'height': screen_height,
            'scaling': scaling_factor,
            'midX': ( screen_width / scaling_factor ) / 2,
            'midY': ( screen_height / scaling_factor ) / 2,
        }
   
    def GetMarkerColors(self):
        return [
            'Orange',
            'Apricot',
            'Yellow',
            'Lime',
            'Olive',
            'Green',
            'Teal',
            'Navy',
            'Blue',
            'Purple',
            'Violet',
            'Pink',
            'Tan',
            'Beige',
            'Brown',
            'Chocolate',
            'none1',
            'none2'
        ]
   
    def DialogSelectMarkerColor(self):
        if self.ui is False or self.disp is False:
            print('Unable to draw window, no UI.')
            return

        dialogWidth = 650
        dialogHeight = 250

        markerColors = []
        markerColorsRows = []
        count = 1
        allColors = self.GetMarkerColors()
        for color in allColors:
            enabled = not color.startswith('none')
            markerColors.append(
                self.ui.CheckBox({
                    "ID": f"MarkerColor_{color}",
                    "Text": color if enabled else '-',
                    "Enabled": enabled,
                    "Weight": 1
                    })
            )
            if ( count % 6 ) == 0 or count == len( allColors ):
                row = len( markerColorsRows )
                markerColorsRows.append(
                    self.ui.HGroup({
                        "ID": f"MarkerColorRow_{row}",
                        "Spacing": 10,
                        "Weight": 1
                        },
                        markerColors
                        )
                )
                markerColors = []
            count = count + 1

        dlg = self.disp.AddWindow(
            {
                "WindowTitle": "MarkerManager (v" + str(self.version) + ")",
                "ID": "DialogSelectColors",
                "Geometry": [ self.screen['midX'], self.screen['midY'], dialogWidth, dialogHeight ],
            },
            [
                self.ui.VGroup({ "Spacing": 10, "Weight": 1 },
                [
                    self.ui.VGap(3),
                    self.ui.Label({ "ID": "MyLabel", "Text": "What colour markers do you want to mark clips with?", "Weight": 0 }),
                    self.ui.VGroup({ "ID": "MarkerColorRows", "Spacing": 0 }, markerColorsRows ),
                    self.ui.VGap(3),
                    self.ui.Button({ "ID": "AcceptButton", "Text": "Let's Go", "Weight": 0 }),
                ]),
            ])
       
        itm = dlg.GetItems()
       
        # The window was closed
        def _closeButton(ev):
            self.disp.ExitLoop()
        dlg.On.DialogSelectColors.Close = _closeButton
 
        def _func(ev):
            checked = []
            for color in self.GetMarkerColors():
                if itm[f"MarkerColor_{color}"].Checked:
                    checked.append( color )
            self.markers = self.GetMarkersByColor( checked )

            dlg.Hide()
           
            if len( self.markers ) == 0:
                self.DialogMessage( "No marker colours were selected.")
            else:
                self.DialogMarkClips()
           
            self.disp.ExitLoop()

        dlg.On.AcceptButton.Clicked = _func
       
        dlg.Show()
        self.disp.RunLoop()
        dlg.Hide()
   
    def DialogMarkClips(self):
        if self.ui is False or self.disp is False:
            print('Unable to draw window, no UI.')
            return
       
        dialogWidth = 650
        dialogHeight = 250

        markerCount = len( self.markers )

        options = {
            "Mark Clips using Multiple Markers" : {
                "description"   : "Markers WITH NAMES will be treated as IN points, the next marker will be treated the OUT point. If the next marker also has a name, it will be an IN point for the next clip.",
                "enabled"       : True,
                "function"      : self.MarkClipsUsingDualMarkers
            },
            "Mark Clips using Marker Duration"  : {
                "description"   : "Marker durations will be used to denote IN and OUT point of clips.",
                "enabled"       : True,
                "function"      : self.MarkClipsUsingMarkerDuration
            },
        }

        firstDescription = ''
        for key, option in options.items():
            firstDescription = option['description']
            break
       
        dlg = self.disp.AddWindow(
            {
                "WindowTitle": "MarkerManager",
                "ID": "DialogMarkClips",
                "Geometry": [ self.screen['midX'], self.screen['midY'], dialogWidth, dialogHeight ],
            },
            [
                self.ui.VGroup({ "Spacing": 20, "Weight": 1 },
                [
                    # Add your GUI elements here:
                    self.ui.Label({ "ID": "MyLabel", "Text": f"Found {markerCount} markers, how should they be used?", "Weight": 0 }),
                    self.ui.ComboBox({ "ID": "MySelector", "Height": 20, "Weight": 0 }),
                    self.ui.TextEdit({ "ID": "Description", "Text": firstDescription, "ReadOnly": True }),
                    self.ui.VGap(4, 1),
                    self.ui.Button({ "ID": "AcceptButton", "Text": "Let's Go", "Weight": 0 }),
                ]),
            ])
       
        itm = dlg.GetItems()

        for key, option in options.items():
            if option['enabled']:
                itm['MySelector'].AddItem( key )
       
        # The window was closed
        def _closeButton(ev):
            self.disp.ExitLoop()
        dlg.On.DialogMarkClips.Close = _closeButton
       
        def _func(ev):
            selected = itm['MySelector'].CurrentText
            itm['Description'].Text = options[ selected ]['description']
            self.markerProcessingFunction = options[ selected ]['function']
        dlg.On.MySelector.CurrentIndexChanged = _func
 
        def _func(ev):
            dlg.Hide()

            if callable( self.markerProcessingFunction ):
                self.markerProcessingFunction()

            clipNo = len(self.clips)
            headers = [
                { 'title': "Index", 'width': 75 },
                { 'title': "Name", 'width': 350 },
                { 'title': "Duration", 'width': 100 },
                { 'title': "Filename", 'width': 150 },
                { 'title': "Color", 'width': 150 },
                { 'title': "In-Point", 'width': 150 },
                { 'title': "Out-Point", 'width': 150 },
                { 'title': "Notes", 'width': 150 },
                { 'title': "", 'width': 150 },
            ]
            rows = []

            for clip in self.clips:
                rows.append([
                    clip['index'],
                    clip['name'],
                    clip['duration'],
                    clip['filename'],
                    clip['color'],
                    self.FramesToDuration( clip['inPoint'] ),
                    self.FramesToDuration( clip['outPoint'] ),
                    clip['note'],
                ])
            
            def _requestRenderLocation(ev):
                self.AskForRenderLocation()

            def _neverMind(ev):
                self.disp.ExitLoop()

            def _addToRenderQueue(ev):
                self.AddClipsToRenderQueue()
                self.disp.ExitLoop()

            buttons = [
                {
                    'events'    : { },
                    'object'    : self.ui.Label({ "ID": "LabelDestination", "Text": "", "Weight":0 })
                # },{
                    # 'events'    : { 'Clicked': _requestRenderLocation },
                    # 'object'    : self.ui.Button({ "ID": "Destination", "Text": 'Choose Location', "Weight": 0 })
                    # 'object'    : self.ui.LineEdit({ "ID": "Destination", "PlaceholderText": "Render Location" })
                },{
                    'events'    : { },
                    'object'    : self.ui.HGap(4,1)
                },{
                    'events'    : { 'Clicked': _neverMind },
                    'object'    : self.ui.Button({ "ID": "Button_NeverMind", "Text": 'Nevermind', "Weight": 0 })
                },{
                    'events'    : { 'Clicked': _addToRenderQueue },
                    'object'    : self.ui.Button({ "ID": "Button_Render", "Text": 'Add to Render Queue', "Weight": 0 })
                }
            ]
           
            self.DialogTreeDisplay( f"Marked {clipNo} clips based on marker positions.", headers, rows, buttons )

            self.disp.ExitLoop()

        dlg.On.AcceptButton.Clicked = _func
       
        dlg.Show()
        self.disp.RunLoop()
        dlg.Hide()

    def DialogTreeDisplay( self, title, headers, rows, buttons ):
        if self.ui is False or self.disp is False:
            print('Unable to draw window, no UI.')
            return
       
        dialogWidth = 750
        dialogHeight = 450

        controls = [
            self.ui.Label({ "ID": "MyTitle", "Text": title, "Weight": 0.25, "Weight": 0  }),
            self.ui.VGap(4),
            self.ui.Tree({ "ID": "MyTree", "SortingEnabled": False }),
            self.ui.VGap(4),
        ]

        if type( buttons ) is list:
            buttonList = []
            for button in buttons:
                if button['object']:
                    buttonList.append( button['object'] )
            controls.append( self.ui.HGroup({ "Spacing": 10, "Weight": 0 }, buttonList ) )
        else:
            controls.append( self.ui.Button({ "ID": "AcceptButton", "Text": 'Okay', "Weight": 0 }) )
       
        dlg = self.disp.AddWindow(
            {
                "WindowTitle": "MarkerManager",
                "ID": "DialogMessage",
                "Geometry": [ self.screen['midX'], self.screen['midY'], dialogWidth, dialogHeight ],
            },
            [
                self.ui.VGroup({ "Spacing": 10 }, controls ),
            ])
       
        itm = dlg.GetItems()

        # Add items to tree
        header = itm['MyTree'].NewItem()
        index = 0
        for h in headers:
            header.Text[index] = h['title']
            index = index + 1
        itm['MyTree'].SetHeaderItem(header)
        itm['MyTree'].ColumnCount = len(headers) -1

        index = 0
        for h in headers:
            itm['MyTree'].ColumnWidth[index] = h['width']
            index = index + 1

        for columns in rows:
            column = itm['MyTree'].NewItem()

            index = 0
            for c in columns:
                column.Text[index] = str(c)
                index = index + 1

            itm['MyTree'].AddTopLevelItem(column)

        # Handle button presses
        for button in buttons:
            try:
                for event, eventFunc in button['events'].items():
                    setattr( dlg.On[ button['object']['ID'] ], event, eventFunc)
            except:
                print('Error with button.')
       
        # The window was closed
        def _closeButton(ev):
            self.disp.ExitLoop()
        dlg.On.DialogMessage.Close = _closeButton

        def _func(ev):
            self.disp.ExitLoop()
        dlg.On.AcceptButton.Clicked = _func
       
        dlg.Show()
        self.disp.RunLoop()
        dlg.Hide()

    def DialogTextDisplay( self, title, text ):
        if self.ui is False or self.disp is False:
            print('Unable to draw window, no UI.')
            return
       
        dialogWidth = 750
        dialogHeight = 450
       
        dlg = self.disp.AddWindow(
            {
                "WindowTitle": "MarkerManager",
                "ID": "DialogMessage",
                "Geometry": [ self.screen['midX'], self.screen['midY'], dialogWidth, dialogHeight ],
            },
            [
                self.ui.VGroup({ "Spacing": 10 },
                [
                    # Add your GUI elements here:
                    self.ui.Label({ "ID": "MyTitle", "Text": title, "Weight": 0.25, "Weight": 0  }),
                    self.ui.VGap(4),
                    self.ui.TextEdit({ "ID": "MyText", "Text": text, "ReadOnly": True, "Weight": 1 }),
                    self.ui.VGap(4),
                    self.ui.Button({ "ID": "AcceptButton", "Text": 'Okay', "Weight": 0 }),
                ]),
            ])
       
        itm = dlg.GetItems()
       
        # The window was closed
        def _closeButton(ev):
            self.disp.ExitLoop()
        dlg.On.DialogMessage.Close = _closeButton

        def _func(ev):
            self.disp.ExitLoop()
        dlg.On.AcceptButton.Clicked = _func
       
        dlg.Show()
        self.disp.RunLoop()
        dlg.Hide()

    def DialogMessage( self, text ):
        if self.ui is False or self.disp is False:
            print('Unable to draw window, no UI.')
            return
       
        dialogWidth = 450
        dialogHeight = 100
       
        dlg = self.disp.AddWindow(
            {
                "WindowTitle": "MarkerManager",
                "ID": "DialogMessage",
                "Geometry": [ self.screen['midX'], self.screen['midY'], dialogWidth, dialogHeight ],
            },
            [
                self.ui.VGroup({ "Spacing": 10, },
                [
                    # Add your GUI elements here:
                    self.ui.Label({ "ID": "MyLabel", "Text": text }),
                    self.ui.Button({ "ID": "AcceptButton", "Text": 'Okay' }),
                ]),
            ])
       
        itm = dlg.GetItems()
       
        # The window was closed
        def _closeButton(ev):
            self.disp.ExitLoop()
        dlg.On.DialogMessage.Close = _closeButton

        def _func(ev):
            self.disp.ExitLoop()
        dlg.On.AcceptButton.Clicked = _func
       
        dlg.Show()
        self.disp.RunLoop()
        dlg.Hide()

    def AddMarker(self, frame, color, name, comment, duration=1, custom_data=None):
        return self.timeline.AddMarker(frame, color, name, comment, duration, custom_data)

    def DeleteMarker(self, frame):
        return self.timeline.DeleteMarkerAtFrame(frame)

    def GetMarkers(self):
        return self.timeline.GetMarkers()

    def GetMarkersByColor(self, color):
        if type(color) != list:
            color = [ color ]
        markers = self.GetMarkers()
        return {frame: details for frame, details in markers.items() if details['color'] in color}
   
    def Markers( self ):
        self.markers = self.GetMarkers()
        return self
   
    def MarkersByColor( self, color ):
        self.markers = self.GetMarkersByColor( color )
        return self

    def EditMarkers(self, markers, color=None, name=None, note=None, duration=None, custom_data=None):
        for frame, marker in markers.items():
            if self.timeline.DeleteMarkerAtFrame(frame):
                self.AddMarker(
                    frame,
                    color if color else marker['color'],
                    name if name else marker['name'],
                    note if note else marker['note'],
                    duration if duration else marker['duration'],
                    custom_data if custom_data else marker['customData']
                )

    def DeleteAllMarkers(self):
        markers = self.GetMarkers()
        for frame in markers.keys():
            self.DeleteMarker(frame)

    def GetSettings(self):
        return self.project.GetSetting()
   
    def CalculateAspectRatio(self, height, width):
        # Convert inputs to integers if they're not already
        height = int(height)
        width = int(width)

        # Import the gcd function from math module
        from math import gcd

        # Calculate the Greatest Common Divisor of height and width
        gcd_val = gcd(height, width)

        # Divide height and width by the gcd to get simplified ratio
        ratio_height = height // gcd_val
        ratio_width = width // gcd_val

        # Return the ratio in the required format
        return f"{ratio_width}_{ratio_height}"
   
    def MarkClip( self, inPoint, outPoint, marker, index ):

        # Duration in frames
        frames = int( outPoint ) - int( inPoint )

        # Duration as timecode
        duration = self.CalculateDuration( inPoint, outPoint )

        # Filename
        fileName = self.SanitizeFilename( marker['name'] )
        fileName = f"{index:02d}_{fileName}"

        # print( f"Marked clip '{fileName}' with duration {duration}")
       
        clip = {
            'index'     : f"{index:002d}",
            'inPoint'   : inPoint,
            'outPoint'  : outPoint,
            'filename'  : fileName,
            'name'      : marker['name'],
            'note'      : marker['note'],
            'color'     : marker['color'],
            'frames'    : frames,
            'duration'  : duration,
        }

        self.clips.append(clip)

    def ListClips( self ):
        total_clips = len( self.clips )
        print( f"Clips marked: {total_clips}" )

        for clip in self.clips:
            print(clip)

    def AddClipsToRenderQueue( self ):
        if self.renderLocation == False:
            self.AskForRenderLocation()

        total_clips = len( self.clips )
        print( f"Clips marked: {total_clips}" )

        for clip in self.clips:
            print(clip)
            self.AddClipToRenderQueue( clip['inPoint'], clip['outPoint'], self.renderLocation, clip['filename'] )

    def AddClipToRenderQueue( self, inPoint, outPoint, location, fileName ):
        frameRate = self.project.GetSetting('timelineFrameRate')
        height = self.project.GetSetting('timelineResolutionHeight')
        width = self.project.GetSetting('timelineResolutionWidth')
        ratio = self.CalculateAspectRatio( height, width )
        inPoint = int( inPoint )
        outPoint = int( outPoint )

        settings = {
            "SelectAllFrames": False,  # Bool (when set True, the settings MarkIn and MarkOut are ignored)
            "MarkIn": inPoint,  # int
            "MarkOut": outPoint,  # int
            "TargetDir": location,  # string
            "CustomName": fileName,  # string
            "UniqueFilenameStyle": 0,  # 0 - Prefix, 1 - Suffix
            "ExportVideo": True,  # Bool
            "ExportAudio": True,  # Bool
            "FormatWidth": width,  # int
            "FormatHeight": height,  # int
            "FrameRate": frameRate,  # float (examples: 23.976, 24)
            "PixelAspectRatio": ratio,  # string (for SD resolution: "16_9" or "4_3") (other resolutions: "square" or "cinemascope")
            "VideoQuality": 0,  # possible values for current codec (if applicable):
                                # 0 (int) - will set quality to automatic
                                # [1 -> MAX] (int) - will set input bit rate
                                # ["Least", "Low", "Medium", "High", "Best"] (String) - will set input quality level
            "AudioCodec": "aac",  # string (example: "aac")
            "AudioBitDepth": 16,  # int
            "AudioSampleRate": 44100,  # int
            "ColorSpaceTag": "Same as Project",  # string (example: "Same as Project", "AstroDesign")
            "GammaTag": "Same as Project",  # string (example: "Same as Project", "ACEScct")
            "ExportAlpha": False,  # Bool
            "EncodingProfile": "Main10",  # string (example: "Main10"). Can only be set for H.264 and H.265.
            "MultiPassEncode": False,  # Bool. Can only be set for H.264.
            "AlphaMode": 0,  # 0 - Premultiplied, 1 - Straight. Can only be set if "ExportAlpha" is true.
            "NetworkOptimization": False  # Bool. Only supported by QuickTime and MP4 formats.
        }

        self.project.SetRenderSettings( settings )

        jobId = self.project.AddRenderJob()
        if jobId:
            print(f"Added render job with id: {jobId}")
        else:
            print("Failed to add render job")

    def Slugify(self, value, allow_unicode=False):
        """
        Taken from https://github.com/django/django/blob/master/django/utils/text.py
        Convert to ASCII if 'allow_unicode' is False. Convert spaces or repeated
        dashes to single dashes. Remove characters that aren't alphanumerics,
        underscores, or hyphens. Convert to lowercase. Also strip leading and
        trailing whitespace, dashes, and underscores.
        """
        value = str(value)
        if allow_unicode:
            value = unicodedata.normalize('NFKC', value)
        else:
            value = unicodedata.normalize('NFKD', value).encode('ascii', 'ignore').decode('ascii')
        value = re.sub(r'[^\w\s-]', '', value.lower())
        return re.sub(r'[-\s]+', '-', value).strip('-_')

    def IsWriteable( self, path ):
        try:
            testfile = tempfile.TemporaryFile(dir = path)
            testfile.close()
        except OSError as e:
            if e.errno == errno.EACCES:  # 13
                return False
            e.filename = path
            raise
        return True

    def AskForRenderLocation( self ):
        self.renderLocation = os.path.normpath( self.fusion.RequestDir() )

        # Add dialog at some point to check render location
        # if os.path.isdir( str(renderLocation) ) and self.IsWriteable( str(renderLocation) ):
        #     return renderLocation
        # else:
        #     print('Unable to render to that directory, please choose another.')
        #     return self.AskForRenderLocation()
       
    def SanitizeFilename( self, string ):
        string = string.split('-')[0]
        return self.Slugify( string )
       
    def CalculateDuration( self, inPoint, outPoint ):
        if int( outPoint ) < int( inPoint ):
            return f'00:00'
        else:
            diff_frames = int( outPoint ) - int( inPoint )
            return self.FramesToDuration( diff_frames )
       
    def FramesToDuration( self, diff_frames ):
            frameRate = self.project.GetSetting('timelineFrameRate')
            diff_seconds, diff_frames = divmod( diff_frames, frameRate )
            diff_minutes, diff_seconds = divmod( diff_seconds, 60 )
            diff_hours, diff_minutes = divmod( diff_minutes, 60 )

            diff_hours = int(diff_hours)
            diff_minutes = int(diff_minutes)
            diff_seconds = int(diff_seconds)
            diff_frames = int(diff_frames)

            return f'{diff_hours:02d}:{diff_minutes:02d}:{diff_seconds:02d}:{diff_frames:02d}'

    def MarkClipsUsingDualMarkers( self, markers = {} ):
        if len( markers ) == 0 and len( self.markers ) != 0:
            markers = self.markers

        index = 1
        markIn = -1
        markOut = -1
        markerIn = False

        # Account for timelines that dont start with 00:00:00:00 timecode
        startFrame = self.timeline.GetStartFrame()

        # Loop through markers and mark clips based on the chosen clip-selection style
        # - At the moment it marks clips using markers with names and the next marker in the sequence as the out point
        for frame, marker in markers.items():

            if not marker['name'].startswith('Marker '):
                if markIn != -1:
                    markOut = startFrame + frame
                else:
                    markIn = startFrame + frame
                    markOut = -1
                    markerIn = marker
            else:
                markOut = startFrame + frame

            if markIn != -1 and markOut != -1 and markerIn:
                self.MarkClip( markIn, markOut, markerIn, index )
                markIn = -1
                markOut = -1
                index = index + 1
                markerIn = False

        return self
    
    def MarkClipsUsingMarkerDuration(self, markers={}):
        if len(markers) == 0 and len(self.markers) != 0:
            markers = self.markers
        
        index = 1
        startFrame = self.timeline.GetStartFrame()
        
        # Loop through markers and use their duration
        for frame, marker in markers.items():

            markIn = startFrame + frame
            # Get marker duration (in frames)
            duration = marker.get('duration', 0)  # Default to 0 if no duration
            if duration > 0:
                markOut = markIn + duration
                self.MarkClip(markIn, markOut, marker, index)
                index += 1
        
        return self

def GetBMD():
    try:
    # The PYTHONPATH needs to be set correctly for this import statement to work.
    # An alternative is to import the DaVinciResolveScript by specifying absolute path (see ExceptionHandler logic)
        import DaVinciResolveScript as bmd
    except ImportError:
        if sys.platform.startswith("darwin"):
            expectedPath="/Library/Application Support/Blackmagic Design/DaVinci Resolve/Developer/Scripting/Modules/"
        elif sys.platform.startswith("win") or sys.platform.startswith("cygwin"):
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
            print("Unable to find module DaVinciResolveScript, there could be a few reasons for this:")
            print("- DaVinciResolveScript is not discoverable by python")
            print("- Python version is higher than 3.12 and doesn't support the imp module (deprecated)")
            print("For a default DaVinci Resolve installation, the module is expected to be located in: "+expectedPath)
            sys.exit()

    return bmd

mm = MarkerManager()

