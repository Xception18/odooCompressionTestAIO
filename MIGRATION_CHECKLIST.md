# UI Modernization - Manual Test Checklist

**Test in BOTH dark and light mode**

## Startup & Navigation
- [ ] App starts without errors
- [ ] All five pages are accessible
- [ ] Window resizes without clipping/overlap
- [ ] Sidebar navigation works
- [ ] Theme toggle switches properly

## Files Page
- [ ] Browse buttons open file dialogs
- [ ] Selected paths display correctly
- [ ] Invalid paths show validation state
- [ ] Sheet name inputs work

## Database Page
- [ ] All DB fields editable
- [ ] Password field has show/hide toggle
- [ ] Test connection: success shows inline result
- [ ] Test connection: failure shows error inline
- [ ] No popup dialogs

## Process Dashboard
- [ ] Pipeline stepper displays
- [ ] Individual action buttons work
- [ ] Full Process runs all steps in order
- [ ] Progress indicators update
- [ ] Success/failure states correct

## Specific Workflows
- [ ] Copy Data → works, shows toast notification
- [ ] Get BJDT ID → works, no request timeout
- [ ] Generate CSV → creates files in Output dir
- [ ] Upload to DB → actual success only
- [ ] Upload duplicate → error shown, NO success message
- [ ] Full Process → steps run sequentially (not parallel)

## Rencana Benda Uji
- [ ] Start button works
- [ ] Stop button terminates gracefully
- [ ] UI stays responsive during processing
- [ ] Status updates appear
- [ ] Button returns to idle state after stop

## ODOO Sync
- [ ] Start sync → daemon starts, indicator updates
- [ ] Stop sync → daemon stops cleanly
- [ ] Status pill matches actual thread state
- [ ] Sidebar indicator agrees with sync state

## Logs Page
- [ ] Log levels have correct colors (INFO/WARNING/ERROR)
- [ ] Filter by level works
- [ ] Search works
- [ ] Clear logs works
- [ ] Auto-scroll toggle works
- [ ] Line cap prevents infinite growth (>5000 lines)
- [ ] Save logs to file works

## Beban Config Editor
- [ ] Opens in matching theme
- [ ] Drag-and-drop reordering works
- [ ] Edit/Save works

## Shutdown
- [ ] Close during idle → closes immediately
- [ ] Close during process → shows confirmation
- [ ] Close after confirm → stops threads cleanly
- [ ] No orphan Chrome/Selenium processes

## Display Scaling (Windows)
- [ ] 100% scaling - text readable, no blur
- [ ] 125% scaling - text readable, no blur
- [ ] 150% scaling - text readable, no blur
- [ ] Multi-monitor with different DPI

## Thread Safety
- [ ] No TclError crashes during processing
- [ ] No UI freezes
- [ ] All progress updates smooth
- [ ] Toasts appear correctly from worker threads

## Regression Checks
- [ ] Excel processing produces same results
- [ ] BJDT API calls work
- [ ] CSV format unchanged
- [ ] DB uploads work
- [ ] Grid Benda Uji launches (separate process)

---

**Bug Fixes Verified:**
- [ ] Bug #1: CSV generated before upload (fixed race)
- [ ] Bug #2: No Tk calls from worker threads
- [ ] Bug #3: Upload success only on actual success
- [ ] Bug #4: Stop doesn't block UI thread
- [ ] Bug #5: No hardcoded DB password
- [ ] Bug #6: Early logs not dropped
- [ ] Bug #7: Log widget capped
- [ ] Bug #8: Clean shutdown stops all threads
- [ ] Bug #9: Requests have timeouts
- [ ] Bug #10: Consistent sync button text

**Date Tested:** _________  
**Tested By:** _________  
**Theme:** Dark / Light  
**OS:** Windows ___  
**Display Scaling:** _____%
