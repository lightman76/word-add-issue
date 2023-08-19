# Minimal reproduction of some Word Online add-in API issues with content controls

Clicking the `Insert Entries into content control` button creates a content control with 4 paragraphs and demonstrates several issues I'm experiencing with Word Online. 
We're using code similar to this to add a bibliography to Word Docs.

In desktop Word (both Windows and macOS), the add-in is creating the content just fine. [Screenshot](https://scrible-maintenance.s3.us-west-2.amazonaws.com/office-test1/BugExample-Word-on-macOS.png)

In Word Online, the add-in is inserting the paragraphs of the content control in reverse order and isn't formatting them. [Screenshot](https://scrible-maintenance.s3.us-west-2.amazonaws.com/office-test1/BugExample-Word-online.png)

<p style="font-weight: bold; margin-top: 10px;">Content Control Insertion Order</p>
<p>In Desktop versions of Office 365 Word, the entries are inserted in the correct increasing order (1,2,3,4).</p>
<p>In Office 365 Online Word, each batch of entries gets inserted in reverse order (4,3,2,1).</p>

<p style="font-weight: bold; margin-top: 10px;">Formatting</p>
<p>In Desktop versions of Office 365 Word, the paragraphs in the content control have a hanging indent and are double spaced as the code instructs.</p>
<p>In Office 365 Online Word, the paragraphs in the content control ignores all formatting instructions applied to the paragraph.</p>

<p style="font-weight: bold; margin-top: 10px;">Content Control cannotDelete property</p>
<p>In Desktop versions of Office 365 Word, the content control can be deleted by the user when the <code>cannotDelete</code> property is set to false.</p>
<p>In Office 365 Online Word, the content control can NOT be deleted by the user regardless of the setting of the <code>cannotDelete</code> property.</p>

