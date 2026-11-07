# Office for the web — observed reference

Captured live on 2026-11-07 from a signed-in Microsoft 365 session (Word for the web,
`word-edit.officeapps.live.com`), using REA (`list_browser_targets`, `inspect_web_page`) for the
host page and CDP accessibility capture for the editor, which runs in a cross-origin iframe that
REA will not attach to. Structure only: no Microsoft code, markup or assets were copied.

## Word home (signed in)
Welcome heading with the user's name · quick actions region · **Create blank document** ·
**Upload a file** · "Create with templates" with a category tablist (Recommended, CVs, Invoices,
"+4 more categories"), a searchable template gallery and a Next pager · "My documents" with the
empty state "Your recent files will show up here / Upload an existing file or try one of the
templates above." · App launcher · Settings · Account manager.

## Word editor chrome
Header: File · Comments · Settings · Share · More Collaboration Options · "Search for tools, help
and more (Option+Q)" · Ribbon Display Options · Mode Menu (Editing) · "Saved to OneDrive".
Status bar: "Page 1 of 1" · "0 words" · language · Editor suggestions · Zoom Out · 100% · Zoom In ·
Switch to light/dark background.

## Ribbon tabs and groups
| Tab | Groups and commands |
| --- | --- |
| Home | Undo/Clipboard (Undo, Clipboard, Format Painter) · Font (Font Name, Font Size, Grow Font Size, Shrink Font Size, Bold, Italic, Underline, Strikethrough, Change Case, Text Highlight Color, Font Color) · Paragraph (Bullets, Numbering, Checklist, Decrease/Increase Indent, Text Alignment, Border Options, Paragraph Shading, Line Spacing, Paragraph Options) · Styles (Normal, No Spacing, Heading 1, Headings and Other Styles) · Editing (Find) · Voice (Dictate) · Proofing (Editor) · Add-ins |
| Insert | Page Break · Table · Picture · Drawing · Link · Table of Contents · Bookmarks · Online Video · New Comment · Header & Footer · Page Numbers · Equation · Symbol · Emoji |
| Layout | Page Setup (Margins, Orientation, Size, Columns, Breaks, Line Numbers) · Paragraph (More/Less indent) · Page Background (Page Borders, Page Color) |
| References | Table of Contents · Insert Footnote · Insert Endnote · Search · Citations |
| Review | Proofing (Editor, Spelling & Grammar, Word Count, Check Accessibility) · Translate · Comments (New, Delete, Previous, Next, Show Comments) · Markup view · Tracking (Track Changes, Accept, Reject, Previous, Next) |
| View | Document Views (Separate Pages, Reading View, Immersive Reader) · Zoom · Show (Ruler, Navigation Pane, Header & Footer, Footnotes, Endnotes) · Dark Mode |

## What Office One takes from this
Mirrored: the six-tab ribbon, the group names and their command grouping, the status-bar contents,
the template-gallery home, and the Find/Replace, page-setup, styles, word-count and navigation-pane
commands. Not implemented: Dictation, Add-ins, Equation, Online Video, Citations, Translate,
real-time Comments/Track Changes, Accessibility Checker, Immersive Reader.
