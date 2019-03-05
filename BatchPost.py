#Author-Yuval Rakavy
#Description-Bathc post all tool path according to Setup/Folder structure (each setup in its own directory) and ecch folder as its own .NC file.

import adsk.core, adsk.fusion, adsk.cam, traceback, adsk
import os, re, time
from functools import reduce

app = adsk.core.Application.get()
if app:
    ui = app.userInterface
handlers = []       # Used to avoid garabge collection of handler classes

class Struct:
    def __init__(self, **v):
        self.__dict__.update(v)

def getCamObject():
    doc = app.activeDocument
    products = doc.products
    product = products.itemByProductType('CAMProductType')
    
    if product == None:
        ui.messageBox('There are no CAM operations in the active document.  This script requires the active document to contain at least one CAM operation.',
                        'No CAM Operations Exist',
                        adsk.core.MessageBoxButtonTypes.OKButtonType,
                        adsk.core.MessageBoxIconTypes.CriticalIconType)
        return None
    else:
        return adsk.cam.CAM.cast(product)


class BatchPostCommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()
        
    def notify(self, args):
        try:
            cam = getCamObject()
            if cam == None:
                return
                
            eventArgs = adsk.core.CommandCreatedEventArgs.cast(args)
            command = eventArgs.command
            
            onExecute = BatchPostCommandExecuteHandler()
            command.execute.add(onExecute)
            onDestory = BatchPostCommandDestoryHandler()
            command.destroy.add(onDestory)
            onInputChanged = BatchPostCommandInputChangedHandler()
            command.inputChanged.add(onInputChanged)
            
            handlers.append(onExecute)
            handlers.append(onDestory)
            handlers.append(onInputChanged)
            
            inputs = command.commandInputs
            allOrSpecificSetup = inputs.addRadioButtonGroupCommandInput('idAllOrSpecificSetup', 'Setup Selection')
            allOrSpecificSetupItems = allOrSpecificSetup.listItems
            allOrSpecificSetupItems.add('All setups', True)
            allOrSpecificSetupItems.add('Select setup', False)
            
            setupsDropbox = inputs.addDropDownCommandInput('idSetups', 'Setup to Post', adsk.core.DropDownStyles.TextListDropDownStyle)
            [setupsDropbox.listItems.add(setup.name, False, '') for setup in cam.setups]

            postProcessorAttribute = cam.attributes.itemByName('BatchPost','PostProcessor')
            postProcessorFullPath = postProcessorAttribute.value if postProcessorAttribute is not None else 'NotSpecified'
            (_, postProcessorFile) = os.path.split(postProcessorFullPath)
            
            selectPostProcessorFile = inputs.addRadioButtonGroupCommandInput('idSelectPostProcessor', 'Post Processor')
            selectPostProcessorFile.listItems.add("Use '%s'" % postProcessorFile, False)
            selectPostProcessorFile.listItems.add('Select another', False)
            
            if postProcessorAttribute is None:
                selectPostProcessorFile.listItems.item(1).isSelected = True # No previous value, force the user to select
                selectPostProcessorFile.isVisible = False
            else:
                selectPostProcessorFile.listItems.item(0).isSelected = True # Default is to use the previous one
                selectPostProcessorFile.isVisible = True
            
            selectOutputFolder = inputs.addRadioButtonGroupCommandInput('idSelectOutputDirectory', 'Post output directory')
            selectOutputFolder.listItems.add('Use previous', False)
            selectOutputFolder.listItems.add('Select another', False)
            
            countDrills = inputs.addBoolValueInput('idCountDrills', 'Count drills', True)
            countDrillsAttribute = cam.attributes.itemByName('BatchPost', 'CountDrils')
            countDrills.value = countDrillsAttribute is None or countDrillsAttribute.value == 'True'
            
            specificSetupNameAttribute = cam.attributes.itemByName('BatchPost', 'UseSetup')
 
            if specificSetupNameAttribute and cam.setups.itemByName(specificSetupNameAttribute.value):
                allOrSpecificSetup.listItems.item(1).isSelected = True
                specificSetupItem = next(item for item in setupsDropbox.listItems if item.name == specificSetupNameAttribute.value)
                specificSetupItem.isSelected = True if specificSetupItem is not None else None
                setupsDropbox.isEnabled = True
            else:
                allOrSpecificSetup.listItems.item(0).isSelected = True
                setupsDropbox.selectedItem = setupsDropbox.name            
                setupsDropbox.isEnabled = False
    
            selectOutputFolder = adsk.core.RadioButtonGroupCommandInput.cast(inputs.itemById('idSelectOutputDirectory'))
            outputDirectoryAttribute = cam.attributes.itemByName('BatchPost', 'OutputDirectory')
            
            if outputDirectoryAttribute is None or not os.path.isdir(outputDirectoryAttribute.value):
                selectOutputFolder.listItems.item(1).isSelected = True
                selectOutputFolder.isVisible = False   # No previous directory, the user must select one
            else:
                selectOutputFolder.listItems.item(0).isSelected = True
                selectOutputFolder.isVisible = True   # There is a previous directory, the user can select to use it or not
                selectOutputFolder.listItems.item(0).name = 'Use ' + ('...' + outputDirectoryAttribute.value[-20:] if len(outputDirectoryAttribute.value) > 20 else outputDirectoryAttribute.value)
                
        except:
            print('BatchPostCommandCreatedHandler Failed:\n{}'.format(traceback.format_exc()))
            

class BatchPostCommandInputChangedHandler(adsk.core.InputChangedEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args):
        eventArgs = adsk.core.InputChangedEventArgs.cast(args)
        changedInput = eventArgs.input

        if changedInput.id == 'idAllOrSpecificSetup':
            allOrSpecificSetup = adsk.core.RadioButtonGroupCommandInput.cast(changedInput)
            inputs = eventArgs.firingEvent.sender.commandInputs
            setupsDropbox = inputs.itemById('idSetups')
            setupsDropbox.isEnabled = allOrSpecificSetup.selectedItem.index != 0
            

class BatchPostCommandDestoryHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()
        
    def notify(self):
        try:
            # when the command is done, terminate the script
            # this will release all globals which will remove all event handlers
            adsk.terminate()
        except:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

        

class BatchPostCommandExecuteHandler(adsk.core.CommandEventHandler):
    def __init__(self):

        super().__init__()
        
    def notify(self, args):
        try:
            cam = getCamObject()
            
            eventArgs = adsk.core.CommandEventArgs.cast(args)
            command = eventArgs.firingEvent.sender
            inputs = command.commandInputs

            allOrSpecificSetup = adsk.core.RadioButtonGroupCommandInput.cast(inputs.itemById('idAllOrSpecificSetup'))
            setupsDropbox = adsk.core.DropDownCommandInput.cast(inputs.itemById('idSetups'))
            
            countDrills = adsk.core.BoolValueCommandInput.cast(inputs.itemById('idCountDrills'))
            cam.attributes.add('BatchPost', 'CountDrills', 'True' if countDrills.value else 'False')

            postThisSpecificSetup = cam.setups.itemByName(setupsDropbox.selectedItem.name) if allOrSpecificSetup.selectedItem.index > 0 else None
            
            if postThisSpecificSetup:
                cam.attributes.add('BatchPost', 'UseSetup', postThisSpecificSetup.name)
            else:
                a = cam.attributes.itemByName('BatchPost', 'UseSetup')
                a.deleteMe() if a is not None else None

            selectPostProcessorFile = adsk.core.RadioButtonGroupCommandInput.cast(inputs.itemById('idSelectPostProcessor'))
            postProcessorAttribute = cam.attributes.itemByName('BatchPost','PostProcessor')

            if postProcessorAttribute is None or selectPostProcessorFile.listItems.item(1).isSelected:
                fileDialog = ui.createFileDialog()
                fileDialog.filter = 'Post processor files (*.cps);;All files (*.*)'
                fileDialog.isMultiSelectEnabled = False
                fileDialog.title = 'Select Post Processor to use'
                fileDialog.initialDirectory = cam.personalPostFolder
                
                if fileDialog.showOpen() == adsk.core.DialogResults.DialogOK:
                    postProcessorFile = fileDialog.filename
                    cam.attributes.add('BatchPost', 'PostProcessor', postProcessorFile)
                else:
                    return
            else:
                postProcessorFile = postProcessorAttribute.value
            
            selectOutputFolder = adsk.core.RadioButtonGroupCommandInput.cast(inputs.itemById('idSelectOutputDirectory'))
            outputDirectoryAttribute = cam.attributes.itemByName('BatchPost', 'OutputDirectory')
            
            if outputDirectoryAttribute is None or selectOutputFolder.listItems.item(1).isSelected or not os.path.isdir(outputDirectoryAttribute.value):
                selectOutputFolderDialog = ui.createFolderDialog()
                selectOutputFolderDialog.title = 'Select root folder for product NC files'
                
                if selectOutputFolderDialog.showDialog() == adsk.core.DialogResults.DialogOK:
                    outputDirectory = selectOutputFolderDialog.folder
                else:
                    return
                    
                cam.attributes.add('BatchPost', 'OutputDirectory', outputDirectory)
            else:
                outputDirectory = outputDirectoryAttribute.value

            progressSteps = reduce((lambda steps, setup: steps + setup.folders.count), [setup for setup in cam.setups if postThisSpecificSetup is None or setup == postThisSpecificSetup], 0)
            progressDialog = ui.createProgressDialog()
            progressDialog.show('Generate Toolpath files', 'Starting...', 0, progressSteps)
            
            batchPostSettings = Struct(
                cam = cam,
                progressDialog = progressDialog,
                outputDirectory = outputDirectory,
                postProcessorFile = postProcessorFile,
                drillsCounter = DrillsCounter() if countDrills.value else None
            )
            
            [self.postSetup(batchPostSettings, setup) for setup in cam.setups if postThisSpecificSetup is None or setup == postThisSpecificSetup]
            progressDialog.hide()

        except:
            print('BatchPostCommandExecutedHandler Failed:\n{}'.format(traceback.format_exc()))
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

    def postSetup(self, batchPostSettings, setup):
        assert type(setup) is adsk.cam.Setup
        
        # pass on all the folder in the setup
        for folderIndex in range(setup.folders.count):
            self.postSetupFolder(batchPostSettings, setup, setup.folders.item(folderIndex), folderIndex)
            
            if batchPostSettings.progressDialog.wasCancelled:
                return
    
    def postSetupFolder(self, batchPostSettings, setup, folder, folderIndex):
        assert type(setup) is adsk.cam.Setup
        assert type(folder) is adsk.cam.CAMFolder
    
        folderName = re.sub(r' \(\d+\)', '', folder.name)
        ncDirectory = os.path.join(batchPostSettings.outputDirectory, setup.name)
        ncProgramName = str(folderIndex) + "_" + folderName
        ncFilename = os.path.join(ncDirectory, ncProgramName + ".nc")
        
        os.makedirs(ncDirectory, exist_ok=True)
        
        needRegeneration = [operation for operation in folder.allOperations if not operation.hasToolpath or not operation.isToolpathValid]
        for operation in needRegeneration:
            batchPostSettings.progressDialog.message = "Generating toolpath for '%s' %s: %s" % (setup.name, folderName, operation.name)
            generationFuture = batchPostSettings.cam.generateToolpath(operation)
            
            while not generationFuture.isGenerationCompleted:
                adsk.core.adsk_doEvents()
        
            if batchPostSettings.progressDialog.wasCancelled:
                return
            
        batchPostSettings.progressDialog.message = "Posting toolpath for '%s' %s" % (setup.name, folderName)
        postInput = adsk.cam.PostProcessInput.create(ncProgramName, batchPostSettings.postProcessorFile, ncDirectory, adsk.cam.PostOutputUnitOptions.DocumentUnitsOutput)
        postInput.isOpenInEditor = False
        
        if not batchPostSettings.cam.postProcess(folder.allOperations, postInput):
            ui.messageBox("Post operation for setup '%s' - %s failed" % (setup.name, folder.name), 'Post operation failed', adsk.core.MessageBoxButtonTypes.OKButtonType, adsk.core.MessageBoxIconTypes.CriticalIconType)
        
        # This delay is needed to ensure that the post is done (otherwise there is a race condition and some post are failing)
        startTime = time.time()
        while time.time() - startTime < 0.5:
            adsk.core.adsk_doEvents()
        
        if batchPostSettings.drillsCounter is not None and "drill" in folderName.lower():
            batchPostSettings.drillsCounter.process(ncFilename)
            
        batchPostSettings.progressDialog.progressValue += 1


def run(context):
    try:
        commandDefs = ui.commandDefinitions
        batchPostCommandDefinition = commandDefs.itemById('BatchPostCommand')

        if not batchPostCommandDefinition: 
            batchPostCommandDefinition = commandDefs.addButtonDefinition('BatchPostCommand', 'Batch Post', 'Batch Post NC')
        
        onCreated = BatchPostCommandCreatedHandler()
        batchPostCommandDefinition.commandCreated.add(onCreated)
        handlers.append(onCreated)
        
        inputs = adsk.core.NamedValues.create()
        batchPostCommandDefinition.execute(inputs)
        adsk.autoTerminate(False)

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

class DrillsCounter:
    
    def count_drills(self, g_code_filename):
        drills = {}
        is_drilling_canned_operation = False
        z_was_above_stock = False

        def is_number(s):
            return s.replace(".", "").replace("-", "").isnumeric()
    
        def add_drill(z):
            z = abs(z)
            drills[z] = drills[z] + 1 if z in drills else 1
    
        def is_z_below_stock(z):
            return z < 0
    
        with open(g_code_filename, 'r') as f:
            for block in f:
                for token in block.split():
                    if token == "G81":
                        is_drilling_canned_operation = True
                    elif token == "G80":
                        is_drilling_canned_operation = False
                    elif token[0] == 'Z' and is_number(token[1:]):
                        z_value = float(token[1:])
    
                        # Count drills by counting the times Z move from being
                        # above stock to being in the stock
                        if not is_drilling_canned_operation:
                            # look for Z changing from negative to positive without
                            # motion in X/Y. This is drilling...
                            if is_z_below_stock(z_value):
                                # if Z was above stock and now it is going down,
                                # it is a drill
                                if z_was_above_stock:
                                    add_drill(z_value)
                                z_was_above_stock = False
                            else:
                                z_was_above_stock = True
    
                if is_drilling_canned_operation:
                    add_drill(z_value)
    
            return drills
        
    def get_adjusted_filename(self, filename, drill_count):
        count_suffix = "_x" + str(drill_count) + ".nc"
        m = re.match("(.*)_x([0-9]*)", filename)
    
        if m is not None:
            base_filename = m.group(1)
        else:
            m = re.match(r'(.*)\.nc', filename)
            base_filename = m.group(1) if m is not None else filename
    
        return base_filename + count_suffix
    
    def get_drill_count(self, drills):
        if len(drills) == 1:
            return next(iter(drills.values()))
        else:
            raise ValueError("No drills in this file" if len(drills) == 0 else "More than one drill depth found in file")
    
    
    def process(self, filename):
        drills = self.count_drills(filename)
    
        try:
            count = self.get_drill_count(drills)
            new_filename = self.get_adjusted_filename(filename, count)
    
            if new_filename != filename:
                os.rename(filename, new_filename)
    
        except ValueError as error:
            ui.messageBox(str(error), 'Drill counting problem', adsk.core.MessageBoxButtonTypes.OKButtonType, adsk.core.MessageBoxIconTypes.WarningIconType)
    
    