# UI Improvement Plan — MSR CV Processing Studio

## Current State
The app uses Streamlit's default components with custom CSS styling (teal `#199E95` theme, Plus Jakarta Sans font). The UI flows: upload → process → split-screen comparison → export.

---

## 1. Empty State & Landing Experience
**Problem:** When no files are uploaded, the page is mostly white with just a title and subtitle.

**Improvements:**
- Add a visually rich empty state with:
  - Animated upload zone (dashed border with pulse animation)
  - Supported format badges (PDF, DOCX, DOC, TXT) as styled pills
  - Key feature highlights in a 3-column grid (Local Parser, GCC Template, Dual Export)
  - Animated hero illustration or icon grid

---

## 2. Enhanced Upload Section
**Problem:** Basic file uploader with minimal feedback.

**Improvements:**
- Add file type badges next to each uploaded file (e.g., `.pdf`, `.docx`)
- Show file size and name in a styled list
- Add a drag-and-drop zone with visual feedback (border color change on hover)
- Show photo upload preview thumbnail if a photo is uploaded
- Add file validation (size limit warning, unsupported format warning)

---

## 3. Processing View
**Problem:** Basic progress bar with minimal visual feedback during batch processing.

**Improvements:**
- Show per-file processing cards with status icons:
  - ⏳ Processing (animated spinner)
  - ✅ Success (green checkmark)
  - ❌ Failed (red X with error tooltip)
- Add a progress step indicator (1. Parse → 2. Extract → 3. Generate DOCX → 4. Generate PDF)
- Show estimated time remaining
- Animate the progress bar with smoother transitions

---

## 4. Results Dashboard
**Problem:** Results are shown in a flat list with a select box. No visual hierarchy.

**Improvements:**
- Replace the select box with **candidate profile cards** (horizontal cards with:
  - Avatar/initials circle with photo if available
  - Candidate name as headline
  - Experience badge (years)
  - Status badge (Success/Error)
  - Quick action buttons (Inspect, Download DOCX, Download PDF)
- For single candidate, show a **summary hero card** with:
  - Name prominently displayed
  - Experience in a large metric
  - Executive summary as a quote-style text block
  - Key details as chip/tag pills (nationality, gender, religion, etc.)

---

## 5. Split-Screen Comparison
**Problem:** Basic two-column layout with plain text areas and markdown.

**Improvements:**
- Make the comparison collapsible with toggle buttons (Raw vs. Parsed view)
- Add syntax highlighting / styled sections for the parsed profile
- Show parsed data as structured cards instead of raw markdown:
  - Contact info chips
  - Experience cards with employer name as header
  - Education row with graduation date
  - Registration badges (PRC, SCFHS)
- Add a "Copy to Clipboard" button for each section

---

## 6. Export Section Redesign
**Problem:** Export buttons are scattered and basic.

**Improvements:**
- Group exports into a dedicated **Export Panel** with:
  - Document preview cards (DOCX, PDF) showing a thumbnail/icon
  - Download buttons with file size info
  - Batch export as a styled card with zip download
  - "Export All" primary button for batch operations
  - Copy profile text button
- Add export format options (e.g., toggle between GCC template vs. clean modern format)

---

## 7. Sidebar Enhancements
**Problem:** Sidebar is basic with just logo, metrics, and status.

**Improvements:**
- Add a **recent candidates** list in the sidebar
- Add session timer / processing stats
- Add a theme toggle (if implementing dark mode)
- Add keyboard shortcut hints (e.g., `Ctrl+P` to process)
- Add version info and links to docs

---

## 8. Animations & Micro-interactions
**Improvements:**
- Fade-in animation for results appearing
- Smooth transitions between states (upload → processing → results)
- Hover effects on cards (subtle lift/shadow)
- Pulse animation on processing button during work
- Success celebration animation on batch completion

---

## 9. Responsive Design
**Improvements:**
- Ensure the layout works well on smaller screens (Streamlit cloud mobile)
- Stack columns vertically on mobile
- Reduce font sizes and spacing for mobile

---

## 10. Accessibility (a11y)
**Improvements:**
- Add `aria-label` attributes to icon-only buttons
- Ensure color contrast meets WCAG AA (teal `#199E95` on white passes)
- Add keyboard navigation support for candidate cards
- Provide text alternatives for all icons/emojis

---

## Priority & Phasing

### Phase 1 — High Impact, Low Effort
| # | Improvement | Estimated Effort |
|---|-------------|-----------------|
| 1 | Empty State & Landing | 2-3 hours |
| 2 | Candidate Profile Cards | 2-3 hours |
| 3 | Processing View Enhancements | 1-2 hours |
| 4 | Sidebar Enhancements | 1-2 hours |

### Phase 2 — Medium Impact, Medium Effort
| # | Improvement | Estimated Effort |
|---|-------------|-----------------|
| 5 | Split-Screen Redesign | 3-4 hours |
| 6 | Export Section Redesign | 2-3 hours |
| 7 | Profile Summary Hero Card | 2-3 hours |

### Phase 3 — Polish
| # | Improvement | Estimated Effort |
|---|-------------|-----------------|
| 8 | Animations & Micro-interactions | 2-3 hours |
| 9 | Responsive Design | 2-3 hours |
| 10 | Accessibility | 1-2 hours |

---

## Technical Notes
- All UI changes are CSS/HTML within `st.markdown(unsafe_allow_html=True)` blocks
- Streamlit components (`st.columns`, `st.container`, `st.expander`) handle layout
- Session state (`st.session_state`) manages interactive state
- No external JS frameworks needed — pure Streamlit + custom CSS
- Font: Keep Plus Jakarta Sans (already loaded via Google Fonts CDN)
- Icons: Use Streamlit's built-in icon support + emoji for simplicity
