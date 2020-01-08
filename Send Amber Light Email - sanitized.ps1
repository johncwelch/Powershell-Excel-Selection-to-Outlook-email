#add the modules needed to draw the UI
#these are for the listbox and the message box
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#initial value for Endscript
$endScript="No"

#check to see if outlook is running. If it isn't, start it
$ProcessActive = Get-Process outlook -ErrorAction SilentlyContinue
if($ProcessActive -eq $null)
{
 $buttonType=[System.Windows.Forms.MessageBoxButtons]::OK

    $dialogTitle="OUtlook Not running"
    $dialogBody="Please make sure outlook is running, then relaunch this script"

    $dialogIcon=[System.Windows.Forms.MessageBoxIcon]::Information

    $endScript=[System.Windows.Forms.MessageBox]::Show($THIS,$dialogBody,$dialogTitle,$buttonType,$dialogIcon)
    exit
}


#the main run loop
Do {

    #instantiate a new form object
    $form = New-Object System.Windows.Forms.Form
    #initial form title
    $form.Text = 'Email Recipient Selection'
    #initial form size (width,height)
    $form.Size = New-Object System.Drawing.Size(300,400)
    #set the minimum size for the form
    $form.MinimumSize = New-Object System.Drawing.Size(300,400)
    #initial form location
    $form.StartPosition = 'CenterScreen'

    #make the OK button
    $OKButton = New-Object System.Windows.Forms.Button
    #location of the button on the form
    $OKButton.Location = New-Object System.Drawing.Point(75,320)
    #initial button size
    $OKButton.Size = New-Object System.Drawing.Size(75,23)

    #ensures the button changes vertical position as the window resizes
    #by not specifiying a horizontal value in Anchor, the button keeps its relative horizontal position
    $OKButton.Anchor='bottom'

    #button text
    $OKButton.Text = 'OK'

    #make this an actual "OK" button
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

    #accepts the click on this button as the OK button 
    $form.AcceptButton = $OKButton

    #bind the button to the form
    $form.Controls.Add($OKButton)

    #same as for OK button
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(150,320)
    $CancelButton.Size = New-Object System.Drawing.Size(75,23)

    #ensures the button changes vertical position as the window resizes
    #by not specifiying a horizontal value in Anchor, the button keeps its relative horizontal position
    $CancelButton.Anchor='bottom'

    $CancelButton.Text = 'Cancel'
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    #this adds a label object within the form
    #the initial bits are about the same as for the buttons
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(280,20)
    $label.Text = 'Select the recipient:'
    $form.Controls.Add($label)

    #now we make the list box that will hold the email addresses
    $listBox = New-Object System.Windows.Forms.ListBox

    #set the initial size, location, and height of the listbox
    $listBox.Location = New-Object System.Drawing.Point(10,40)
    $listBox.Size = New-Object System.Drawing.Size(260,270)
    $listBox.Height = 270

    #this ensures the listbox area resizes as the window resizes
    $listBox.Anchor='top,bottom,left,right'

$listBox.SelectionMode = 'MultiExtended'

    [void] $listBox.Items.Add('email@address.one')
    [void] $listBox.Items.Add(' ')
    [void] $listBox.Items.Add('email@address.two')
    [void] $listBox.Items.Add(' ')

    $form.Controls.Add($listBox)

    $form.TopLevel = $true
    $form.Topmost = $true

$result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $theRecipients = $listBox.SelectedItems
    } else {
        exit
    }

    #set up the first part of the email. Note that for this, `n is the newline character
    $theBodyText="Hi all,`n`nOn our walkarounds, we found the equipment listed below with an amber/red light and/or audio alarm of some kind indicating a possible fault. If you need us to take any action to assist you with managing this problem, please do not hesitate to let us know.`n`n"
    #write-host $theBodyText

    #get the running copy of outlook
    $theOutlookApp=[Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")

    #create a new message with the specified destination email address/subject/
    #note that 0 is not a counter, but type. 0 is message
    $theNewEmail=$theOutlookApp.createitem(0)


    foreach ($emailAddress in $theRecipients)
        {
            $theNewEmail.Recipients.add($emailAddress)|out-null
        }

    $theNewEmail.Subject="Weekly Amber Light Walkthrough Results"

    #get the running copy of excel
    $theExcelApp=[Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")

    #get the selected cells in the current workbook
    $theSelectedCells = $theExcelApp.Selection

    #copy the selection to the clipboard
    $theSelectedCells.Copy()

    #get the actual message body object
    $theWordDoc=$theNewEmail.Getinspector.WordEditor


    #get the message body "selection". If there's no actual selection, this gets the current cursor position
    #in a brand new outlook message, this is at the top of the body over the signature (if any)
    $theOutlookSelection=$theWordDoc.Application.Selection

    #insert the body text at the "beginning" of the cursor position
    $theOutlookSelection.InsertBefore($theBodyText)

    #move the cursor to the end of the message body. This does not move to the end of the sig, it stays above that
    #note that if you don't do this, what you just inserted is "selected", so anything new you insert obliterates
    #the existing body
    $theOutlookSelection.EndOf()

    #get the content of the clipboard as HTML. This makes those cells into an HTML table
    $theCopiedCells = Get-Clipboard -TextFormatType Html

    #insert the html content into the message where we want it to be.
    $theOutlookSelection.Paste($theCopiedCells)

    #display the new message for the hu-mon
    $theNewEmail.Display()


    #check to see if we still want to do this.
    #if they click anything but "yes", we dump out.
    $buttonType=[System.Windows.Forms.MessageBoxButtons]::YesNo

    $dialogTitle="End this script?"
    $dialogBody="Are you sure you want to end this script?"

    $dialogIcon=[System.Windows.Forms.MessageBoxIcon]::Information

    $endScript=[System.Windows.Forms.MessageBox]::Show($THIS,$dialogBody,$dialogTitle,$buttonType,$dialogIcon)

} While ($endScript -eq "No")