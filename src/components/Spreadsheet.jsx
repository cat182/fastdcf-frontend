import React, { useEffect, useState, useRef, useLayoutEffect } from 'react';
import {ascendingOrder, isPercentage, isFirefox, numberToExcelColumn, isDigit, excelColumnToNumber, linearTrend} from "../functions/util"
import {computeCell} from "../functions/formulaParser"
import {EmptyBlock, HeaderRowCell, HeaderColumnCell, Cell, Row} from "./Cells"
import {Selection, CopySelection, Extension} from "./Selections"
import CellEditor from "./CellEditor"
import {ContextMenu, ContextMenuOption} from './ContextMenu';
import {addCommas} from '../functions/util'

let previousEntriesStringified
let previousComputedCells

function computeCells(entries, rowCount, columnCount) {
    if (entries === previousEntriesStringified) { // Memoization
        return previousComputedCells
    }

    const computedCells = []

    for(let i=1; i <= rowCount; i++) {
        for(let j=1; j <= columnCount; j++) {
            computeCell(entries, i, j, computedCells, [], [])
        }
    }

    previousEntriesStringified = entries
    previousComputedCells = computedCells

    return computedCells
}

export default function Spreadsheet({file, onChange, width, height}) {
    if (file === undefined) { file = {} }
    if (file.defaultColumnWidth === undefined) { file.defaultColumnWidth = 80 }
    if (file.defaultRowHeight === undefined) { file.defaultRowHeight = 20 }
    if (file.columnWidths === undefined) { file.columnWidths = [] }
    if (file.rowHeights === undefined) { file.rowHeights = [] }
    if (file.inputs === undefined) { file.inputs = [] }
    if (file.columnCount === undefined) { file.columnCount = 10 }
    if (file.rowCount === undefined) { file.rowCount = 30 }
    
    const [contextData, setContextData] = useState({})
    const [copyData, setCopyData] = useState(null)
    const [selectedCell, setSelectedCell] = useState({row: 1, column: 1})
    const [selectedCells, setSelectedCells] = useState([])
    const [currentSelection, setCurrentSelection] = useState(null)
    const [cellEditorData, setCellEditorData] = useState({position: null, value: null})
    const [widthResizerPosition, setWidthResizerPosition] = useState(null)
    const [heightResizerPosition, setHeightResizerPosition] = useState(null)
    const [keyPressed, setKeyPressed] = useState({ctrl: false, shift: false})
    const [size, setSize] = useState([0, 0]);
    const [selectionRects, setSelectionRects] = useState([])
    //const [extending, setExtending] = useState(false)
    const [selecting, setSelecting] = useState(false)
    const [extensionData, setExtensionData] = useState(null)
    //const [extensionRect, setExtensionRect] = useState(null)

    const scrollRef = useRef(null)
    const mainRef = useRef(null)
    const rowsRef = useRef(null);
    const selectionsRef = useRef(null);
    const contextMenuRef = useRef(null)

    var edgeSize = 60;
    var timer = null;

    //window.addEventListener( "mousemove", handleMousemove, false );

    //drawGridLines();

    // --------------------------------------------------------------------------- //
    // --------------------------------------------------------------------------- //

    // I adjust the window scrolling in response to the given mousemove event.
    function handleMousemove( event ) {
        if (!selecting && !(extensionData && extensionData.extending)) return

        // NOTE: Much of the information here, with regard to document dimensions,
        // viewport dimensions, and window scrolling is derived from JavaScript.info.
        // I am consuming it here primarily as NOTE TO SELF.
        // --
        // Read More: https://javascript.info/size-and-scroll-window
        // --
        // CAUTION: The viewport and document dimensions can all be CACHED and then
        // recalculated on window-resize events (for the most part). I am keeping it
        // all here in the mousemove event handler to remove as many of the moving
        // parts as possible and keep the demo as simple as possible.

        // Get the viewport-relative coordinates of the mousemove event.
        var viewportX = event.clientX;
        var viewportY = event.clientY;

        // Get the viewport dimensions.
        var viewportWidth = scrollRef.current.clientWidth;
        var viewportHeight = scrollRef.current.clientHeight;

        // Next, we need to determine if the mouse is within the "edge" of the
        // viewport, which may require scrolling the window. To do this, we need to
        // calculate the boundaries of the edge in the viewport (these coordinates
        // are relative to the viewport grid system).
        var edgeTop = edgeSize;
        var edgeLeft = edgeSize;
        var edgeBottom = ( viewportHeight - edgeSize );
        var edgeRight = ( viewportWidth - edgeSize );

        var isInLeftEdge = ( viewportX < edgeLeft );
        var isInRightEdge = ( viewportX > edgeRight );
        var isInTopEdge = ( viewportY < edgeTop );
        var isInBottomEdge = ( viewportY > edgeBottom );

        // If the mouse is not in the viewport edge, there's no need to calculate
        // anything else.
        if ( ! ( isInLeftEdge || isInRightEdge || isInTopEdge || isInBottomEdge ) ) {

            clearTimeout( timer );
            return;

        }

        // If we made it this far, the user's mouse is located within the edge of the
        // viewport. As such, we need to check to see if scrolling needs to be done.

        // Get the document dimensions.
        // --
        // NOTE: The various property reads here are for cross-browser compatibility
        // as outlined in the JavaScript.info site (link provided above).
        var documentWidth = Math.max(
            scrollRef.current.scrollWidth,
            scrollRef.current.offsetWidth,
            scrollRef.current.clientWidth
        );
        var documentHeight = Math.max(
            scrollRef.current.scrollHeight,
            scrollRef.current.offsetHeight,
            scrollRef.current.clientHeight
        );

        // Calculate the maximum scroll offset in each direction. Since you can only
        // scroll the overflow portion of the document, the maximum represents the
        // length of the document that is NOT in the viewport.
        var maxScrollX = ( documentWidth - viewportWidth );
        var maxScrollY = ( documentHeight - viewportHeight );

        // As we examine the mousemove event, we want to adjust the window scroll in
        // immediate response to the event; but, we also want to continue adjusting
        // the window scroll if the user rests their mouse in the edge boundary. To
        // do this, we'll invoke the adjustment logic immediately. Then, we'll setup
        // a timer that continues to invoke the adjustment logic while the window can
        // still be scrolled in a particular direction.
        // --
        // NOTE: There are probably better ways to handle the ongoing animation
        // check. But, the point of this demo is really about the math logic, not so
        // much about the interval logic.
        (function checkForWindowScroll() {

            clearTimeout( timer );

            if ( adjustWindowScroll() ) {

                timer = setTimeout( checkForWindowScroll, 30 );

            }

        })();

        // Adjust the window scroll based on the user's mouse position. Returns True
        // or False depending on whether or not the window scroll was changed.
        function adjustWindowScroll() {
            if (!selecting && !(extensionData && extensionData.extending)) return

            // Get the current scroll position of the document.
            var currentScrollX = scrollRef.current.scrollLeft;
            var currentScrollY = scrollRef.current.scrollTop;

            // Determine if the window can be scrolled in any particular direction.
            var canScrollUp = ( currentScrollY > 0 );
            var canScrollDown = ( currentScrollY < maxScrollY );
            var canScrollLeft = ( currentScrollX > 0 );
            var canScrollRight = ( currentScrollX < maxScrollX );

            // Since we can potentially scroll in two directions at the same time,
            // let's keep track of the next scroll, starting with the current scroll.
            // Each of these values can then be adjusted independently in the logic
            // below.
            var nextScrollX = currentScrollX;
            var nextScrollY = currentScrollY;

            // As we examine the mouse position within the edge, we want to make the
            // incremental scroll changes more "intense" the closer that the user
            // gets the viewport edge. As such, we'll calculate the percentage that
            // the user has made it "through the edge" when calculating the delta.
            // Then, that use that percentage to back-off from the "max" step value.
            var maxStep = 50;

            // Should we scroll left?
            if ( isInLeftEdge && canScrollLeft ) {

                var intensity = ( ( edgeLeft - viewportX ) / edgeSize );

                nextScrollX = ( nextScrollX - ( maxStep * intensity ) );

            // Should we scroll right?
            } else if ( isInRightEdge && canScrollRight ) {

                var intensity = ( ( viewportX - edgeRight ) / edgeSize );

                nextScrollX = ( nextScrollX + ( maxStep * intensity ) );

            }

            // Should we scroll up?
            if ( isInTopEdge && canScrollUp ) {

                var intensity = ( ( edgeTop - viewportY ) / edgeSize );

                nextScrollY = ( nextScrollY - ( maxStep * intensity ) );

            // Should we scroll down?
            } else if ( isInBottomEdge && canScrollDown ) {

                var intensity = ( ( viewportY - edgeBottom ) / edgeSize );

                nextScrollY = ( nextScrollY + ( maxStep * intensity ) );

            }

            // Sanitize invalid maximums. An invalid scroll offset won't break the
            // subsequent .scrollTo() call; however, it will make it harder to
            // determine if the .scrollTo() method should have been called in the
            // first place.
            nextScrollX = Math.max( 0, Math.min( maxScrollX, nextScrollX ) );
            nextScrollY = Math.max( 0, Math.min( maxScrollY, nextScrollY ) );

            if (
                ( nextScrollX !== currentScrollX ) ||
                ( nextScrollY !== currentScrollY )
                ) {

                scrollRef.current.scrollTo( nextScrollX, nextScrollY );
                return( true );

            } else {

                return( false );

            }

        }

    }

    useLayoutEffect(() => {
        if(contextData.posX + contextMenuRef.current?.offsetWidth > window.innerWidth){
            setContextData({ ...contextData, posX: -(contextMenuRef.current?.offsetWidth - window.innerWidth)})
        }
        if(contextData.posY + contextMenuRef.current?.offsetHeight > window.innerHeight){
            setContextData({ ...contextData, posY: -(contextMenuRef.current?.offsetHeight - window.innerHeight)})
        }
    }, [contextData])

    useEffect(() => {

            const newSelectionRects = []
            if (selectedCells && !(selectedCells.length === 1 && selectedCells[0].startColumn === selectedCells[0].endColumn && selectedCells[0].startRow === selectedCells[0].endRow)) {
                selectedCells.forEach((selection, i) => {
                    const [startRow, endRow] = ascendingOrder(selection.startRow, selection.endRow)
                    const [startColumn, endColumn] = ascendingOrder(selection.startColumn, selection.endColumn)

                    if (startRow === -1 && startColumn === -1) return
        
                    // Compute left and top distance
                    const rowsRect = rowsRef.current.getBoundingClientRect()
                    const topDistance = rowsRef.current.children[startRow].getBoundingClientRect().top - rowsRect.top - (startRow === 0 ? 0 : 8)
                    const leftDistance = rowsRef.current.children[startRow].children[startColumn].getBoundingClientRect().left - rowsRect.left
                    const endCellRect = rowsRef.current.children[endRow].children[endColumn].getBoundingClientRect()
                    const width = endCellRect.left - leftDistance + endCellRect.width - rowsRect.left
                    const endRowRect = rowsRef.current.children[endRow].getBoundingClientRect()
                    const height = endRowRect.top - topDistance + endRowRect.height - rowsRect.top - (endRow === file.rowCount && endRow !== 0 ? 0 : 8)
        
                    newSelectionRects.push({width: width, height: height, top: topDistance, left: leftDistance})
                })
            }
    
            setSelectionRects(newSelectionRects)

            if (copyData && copyData.selection && !copyData.hidden) {
                const copySelection = copyData.selection
                let [startRow, endRow] = ascendingOrder(copySelection.startRow, copySelection.endRow)
                startRow = Math.max(1, startRow)
                let [startColumn, endColumn] = ascendingOrder(copySelection.startColumn, copySelection.endColumn)
                startColumn = Math.max(1, startColumn)

                const rowsRect = rowsRef.current.getBoundingClientRect()
                const topDistance = rowsRef.current.children[startRow].getBoundingClientRect().top - rowsRect.top - (startRow === 0 ? 0 : 8)
                const leftDistance = rowsRef.current.children[startRow].children[startColumn].getBoundingClientRect().left - rowsRect.left
                const endCellRect = rowsRef.current.children[endRow].children[endColumn].getBoundingClientRect()
                const width = endCellRect.left - leftDistance + endCellRect.width - rowsRect.left
                const endRowRect = rowsRef.current.children[endRow].getBoundingClientRect()
                const height = endRowRect.top - topDistance + endRowRect.height - rowsRect.top - (endRow === file.rowCount ? 0 : 8)

                setCopyData({...copyData, selectionRect: {width: width, height: height, top: topDistance, left: leftDistance}})
            }

            if (extensionData && extensionData.extension) {
                let startRow = extensionData.extension.startRow
                let endRow = extensionData.extension.endRow
                let startColumn = extensionData.extension.startColumn
                let endColumn = extensionData.extension.endColumn

                const rowsRect = rowsRef.current.getBoundingClientRect()
                const topDistance = rowsRef.current.children[startRow].getBoundingClientRect().top - rowsRect.top - (startRow === 0 ? 0 : 8)
                const leftDistance = rowsRef.current.children[startRow].children[startColumn].getBoundingClientRect().left - rowsRect.left
                const endCellRect = rowsRef.current.children[endRow].children[endColumn].getBoundingClientRect()
                const width = endCellRect.left - leftDistance + endCellRect.width - rowsRect.left
                const endRowRect = rowsRef.current.children[endRow].getBoundingClientRect()
                const height = endRowRect.top - topDistance + endRowRect.height - rowsRect.top - (endRow === file.rowCount ? 0 : 8)

                setExtensionData({...extensionData, rect: {width: width, height: height, top: topDistance, left: leftDistance}})
                //setExtensionData((prevState) => {return {...prevState, width: width, height: height, top: topDistance, left: leftDistance}})
            }

    }, [selectedCells, copyData?.selection, file.columnWidths, extensionData?.extension, size])

    useEffect(() => {
        document.addEventListener("mouseup", onMouseUp)
        document.addEventListener("mousedown", onMouseDown)

        return () => {
            document.removeEventListener("mouseup", onMouseUp)
            document.removeEventListener("mousedown", onMouseDown)
        }
    }, [selectedCells, extensionData])

    useEffect(() => {
        scrollRef.current.addEventListener( "mousemove", handleMousemove, false );

        return () => {
            if (!scrollRef.current) return
            scrollRef.current.removeEventListener( "mousemove", handleMousemove, false );
        }
    }, [selecting, extensionData])

    useEffect(() => {
        const keyDownHandler = event => {
            if (event.key === "Control") {
                setKeyPressed({...keyPressed, ctrl: true})
            } else if (event.key === "Shift") {
                setKeyPressed({...keyPressed, shift: true})
            }
        };

        const keyUpHandler = event => {
            if (event.key === "Control") {
                setKeyPressed({...keyPressed, ctrl: false})
            } else if (event.key === "Shift") {
                setKeyPressed({...keyPressed, shift: false})
            }
        };
    
        document.addEventListener('keydown', keyDownHandler);
        document.addEventListener('keyup', keyUpHandler);
    
        return () => {
            document.removeEventListener('keydown', keyDownHandler);
            document.removeEventListener('keyup', keyUpHandler);
        };
      }, []);    

    // Rerender when zoom changes
    useEffect(() => {
        function updateSize() {
          setSize([window.innerWidth, window.innerHeight]);
        }
        window.addEventListener('resize', updateSize);
        updateSize();
        return () => window.removeEventListener('resize', updateSize);
      }, []);

    // Handle cell overflow
    useEffect(() => {
        /*const cellsWithoutRightBorder = []

        for(let i=1; i<rowsRef.current.children.length; i++){
            let row = rowsRef.current.children[i];
            cellsWithoutRightBorder[i] = []
            for(let j=1; j<row.children.length; j++){
                let cell = row.children[j];
                let span = cell.querySelector("span")
                if (!span) continue
                let rect = span.getBoundingClientRect()
                let remainingWidth = rect.width - getColumnWidth(j)
                let newWidth = getColumnWidth(j)
                let nextCellEmpty = getCellValue(i, j + 1) === ""
                let index = j + 1
                while(remainingWidth > 0 && nextCellEmpty) {
                    newWidth += getColumnWidth(index) + 1
                    remainingWidth -= getColumnWidth(index)
                    nextCellEmpty = getCellValue(i, index + 1) === ""
                    if ((remainingWidth > 0 && nextCellEmpty)) {
                        cellsWithoutRightBorder[i][index] = true
                    }
                    index += 1
                }
                if (getCellValue(i, j) === "") {
                    span.parentNode.style.minWidth = "0px"
                    span.parentNode.style.maxWidth = "0px"
                } else {
                    span.parentNode.style.minWidth = newWidth.toString()+"px"
                    span.parentNode.style.maxWidth = newWidth.toString()+"px"
                }
            }
        }

        for(let i=1; i<rowsRef.current.children.length; i++){
            let row = rowsRef.current.children[i];
            for(let j=1; j<row.children.length; j++){
                let cell = row.children[j];
                if (cellsWithoutRightBorder[i][j]) {
                    cell.style.borderRight = "none"
                    cell.style.paddingRight = "1px"
                } else {
                    cell.style.borderRight = null
                    cell.style.paddingRight = null
                }
            }
        }*/
    }, [file]);

    function getLastSelection() {
        for(let i=selectedCells.length-1; i >= 0; i--) {
            if (!(selectedCells[i].startColumn === -1 && selectedCells[i].startRow === -1)) {
                return selectedCells[i]
            }
        }
        
        return null
    }

      //////////////////
     // CONTEXT MENU //
    //////////////////
    function showContextMenu(event) {
        setContextData({ visible: true, posX: event.clientX, posY: event.clientY  })
        document.body.style.cursor = "default"
    }

    function hideContextMenu() {
        setContextData({ ...contextData, visible: false })
        document.body.style.cursor = "cell"
    }

    function onContextMenuCut() {
        if (!isFirefox()) {
            navigator.permissions.query({ name: "clipboard-write" }).then((result) => {
                if (result.state === "granted" || result.state === "prompt") {
                    const lastSelection = getLastSelection()

                    let clipboardText = ""
                    let [startRow, endRow] = ascendingOrder(lastSelection.startRow, lastSelection.endRow)
                    startRow = Math.max(1, startRow)
                    let [startColumn, endColumn] = ascendingOrder(lastSelection.startColumn, lastSelection.endColumn)
                    startColumn = Math.max(1, startColumn)
                    
                    for(let i=startRow; i<=endRow; i++) {
                        for(let j=startColumn; j<=endColumn; j++) {
                            const cellValue = getCellValue(i, j)
                            clipboardText += '"' + cellValue.replaceAll('"', '""') + '"'
                            if (j !== endColumn) {
                                clipboardText += String.fromCharCode(9)
                            }
                        }
                        if (i !== endRow) {
                            clipboardText += String.fromCharCode(13)
                        }
                    }

                    navigator.clipboard.writeText(clipboardText).then(
                        () => {
                            /* clipboard successfully set */
                            setCopyData((prevState) => {return {...prevState, clipText: clipboardText, cut: true}})
                            hideContextMenu()
                        },
                        () => {
                            /* clipboard write failed */
                            console.error("clipboard write failed !")
                        },
                    );
                } else {
                    console.error(result)
                    hideContextMenu()
                }
            });
        } else {
            hideContextMenu()
        }

        const lastSelection = getLastSelection()

        const copyEntries = []
        let [startRow, endRow] = ascendingOrder(lastSelection.startRow, lastSelection.endRow)
        startRow = Math.max(1, startRow)
        let [startColumn, endColumn] = ascendingOrder(lastSelection.startColumn, lastSelection.endColumn)
        startColumn = Math.max(1, startColumn)
        
        for(let i=startRow; i<=endRow; i++) {
            for(let j=startColumn; j<=endColumn; j++) {
                const cellValue = getCellValue(i, j)
                copyEntries[i-startRow] = copyEntries[i-startRow] || []
                copyEntries[i-startRow][j-startColumn] = cellValue
            }
        }
        setCopyData((prevState) => {return {...prevState, selection: lastSelection, entries: copyEntries, cut: true}})
    }

    function onContextMenuCopy() {
        if (!isFirefox()) {
            navigator.permissions.query({ name: "clipboard-write" }).then((result) => {
                if (result.state === "granted" || result.state === "prompt") {
                    const lastSelection = getLastSelection()

                    let clipboardText = ""
                    let [startRow, endRow] = ascendingOrder(lastSelection.startRow, lastSelection.endRow)
                    startRow = Math.max(1, startRow)
                    let [startColumn, endColumn] = ascendingOrder(Math.max(1, lastSelection.startColumn), lastSelection.endColumn)
                    startColumn = Math.max(1, startColumn)
                    for(let i=startRow; i<=endRow; i++) {
                        for(let j=startColumn; j<=endColumn; j++) {
                            const cellValue = getCellValue(i, j)
                            clipboardText += '"' + cellValue.replaceAll('"', '""') + '"'
                            if (j !== endColumn) {
                                clipboardText += String.fromCharCode(9)
                            }
                        }
                        if (i !== endRow) {
                            clipboardText += String.fromCharCode(13)
                        }
                    }

                    navigator.clipboard.writeText(clipboardText).then(
                        () => {
                            /* clipboard successfully set */
                            setCopyData((prevState) => {return {...prevState, clipText: clipboardText, cut: false}})
                            hideContextMenu()
                        },
                        () => {
                            /* clipboard write failed */
                            console.error("clipboard write failed !")
                        },
                    );
                } else {
                    console.error(result)
                    hideContextMenu()
                }
            });
        } else {
            hideContextMenu()
        }

        const lastSelection = getLastSelection()

        const copyEntries = []

        let [startRow, endRow] = ascendingOrder(lastSelection.startRow, lastSelection.endRow)
        startRow = Math.max(1, startRow)
        let [startColumn, endColumn] = ascendingOrder(lastSelection.startColumn, lastSelection.endColumn)
        startColumn = Math.max(1, startColumn)

        for(let i=startRow; i<=endRow; i++) {
            for(let j=startColumn; j<=endColumn; j++) {
                const cellValue = getCellValue(i, j)
                copyEntries[i-startRow] = copyEntries[i-startRow] || []
                copyEntries[i-startRow][j-startColumn] = cellValue
            }
        }
        setCopyData((prevState) => {return {...prevState, selection: lastSelection, entries: copyEntries, cut: false}})
    }

    function incrementReferences(expression, rowIncrement, columnIncrement) {
        let pos = 0
        let quotes = false
        let fixedRowReference = false
        let fixedColumnReference = false
        let rowReferenceStart = undefined
        let columnReferenceStart = undefined
        let columnStarted = false
        
        while (pos < expression.length) {
            const char = expression[pos]
            const nextChar = expression[pos+1]

            if (!quotes && char === '"') {
                quotes = true;
            } else if (quotes && char === '"') {
                if (nextChar && nextChar === '"') {
                    pos += 2;
                    continue;
                } else {
                    quotes = false;
                }
            }

            if (!quotes) {
                /*console.log("---------------")
                console.log(pos)
                console.log(char)
                console.log(columnReferenceStart)
                console.log(rowReferenceStart)*/
                if ((char.match(/[a-z]/i) || char === "$") && (!columnReferenceStart || !columnStarted)) {
                    if (columnReferenceStart === undefined) {
                        columnReferenceStart = pos
                    }
                    if (char === "$") {
                        fixedColumnReference = true
                        pos += 1
                        continue
                    } else {
                        columnStarted = true
                    }
                } else if ((isDigit(char) || (char === "$")) && columnStarted && rowReferenceStart === undefined) {
                    rowReferenceStart = pos
                    if (char === "$") {
                        fixedRowReference = true
                        pos += 1
                        continue
                    }
                } else if (!(isDigit(char) || char.match(/[a-z]/i) || (char === "$"))) {
                    if (rowReferenceStart !== undefined) {
                        let column = expression.substring(columnReferenceStart, rowReferenceStart).replaceAll("$", "")
                        let row = expression.substring(rowReferenceStart, pos).replaceAll("$", "")

                        const initialLength = expression.substring(columnReferenceStart, pos).length
                        column = excelColumnToNumber(column)
                        row = parseInt(row)
                        if (!isNaN(row) && !isNaN(column)) {
                            if (!fixedRowReference) {
                                row += rowIncrement
                            }
                            if (!fixedColumnReference) {
                                column += columnIncrement
                            }

                            row = row.toString()
                            column = numberToExcelColumn(column)
                            const newReference = (fixedColumnReference ? "$" : "") + column + (fixedRowReference ? "$" : "") + row
                            expression = expression.substring(0, columnReferenceStart) + newReference + expression.substring(pos)
                            pos += newReference.length - initialLength - 1
                        }
                    }
                    fixedRowReference = false
                    fixedColumnReference = false
                    rowReferenceStart = undefined 
                    columnReferenceStart = undefined
                    columnStarted = false
                    pos +=1
                    continue
                }

                if (rowReferenceStart != undefined && pos === expression.length-1) {
                    let column = expression.substring(columnReferenceStart, rowReferenceStart)
                    let row = expression.substring(rowReferenceStart, pos+1)

                    column = excelColumnToNumber(column)
                    row = parseInt(row)
                    if (!isNaN(row) && !isNaN(column)) {
                        if (!fixedRowReference) {
                            row += rowIncrement
                        }
                        if (!fixedColumnReference) {
                            column += columnIncrement
                        }

                        row = row.toString()
                        column = numberToExcelColumn(column)
                        const newReference = (fixedColumnReference ? "$" : "") + column + (fixedRowReference ? "$" : "") + row
                        expression = expression.substring(0, columnReferenceStart) + newReference + expression.substring(pos+1)
                    }
                    break
                }
            } else {
                fixedRowReference = false
                fixedColumnReference = false
                rowReferenceStart = undefined 
                columnReferenceStart = undefined
                columnStarted = false
            }

            pos += 1
        }

        return expression
    }

    function onContextMenuPaste() {
        if (!selectedCell) return

        const newEntries = [...file.inputs]
        let cut
        let copyEntries
        let selection
        let clipboardText
        let clipText
        let startRow, endRow
        let startColumn, endColumn

        if (copyData) {
            let sameEntries = true
            cut = copyData.cut
            if (copyData.entries) {
                copyEntries = JSON.parse(JSON.stringify(copyData.entries))
            }
            selection = copyData.selection
            clipboardText = copyData.clipText || null
            if (selection) {
                let result = ascendingOrder(selection.startRow, selection.endRow)
                startRow = Math.max(1, result[0])
                endRow = result[1]
                result = ascendingOrder(selection.startColumn, selection.endColumn)
                startColumn = Math.max(1, result[0])
                endColumn = result[1]
            }
        }

        function pasteFromState(destinationRow, destinationColumn) {
            if (!sameEntries && cut) {hideContextMenu(); return}
            
            if (copyData) {
                if (cut && sameEntries) {
                    for(let i=startRow; i<=endRow; i++) {
                        for(let j=startColumn; j<=endColumn; j++) {
                            if (newEntries[i]) newEntries[i][j] = ''
                        }
                    }
                }

                for(let i=0; i<copyEntries.length; i++) {
                    for (let j=0; j<copyEntries[i].length; j++) {
                        newEntries[i + destinationRow] = newEntries[i + destinationRow] || []
                        newEntries[i + destinationRow][j + destinationColumn] = copyEntries[i][j]
                    }
                }
                
                if (!cut || sameEntries) {
                    setSelectedCells([{startRow: destinationRow, endRow: Math.min(destinationRow + endRow - startRow, file.rowCount), startColumn: destinationColumn, endColumn: Math.min(destinationColumn + endColumn - startColumn, file.columnCount)}])
                }
            }
        }

        function paste(destinationRow, destinationColumn) {
            if (isFirefox()) {pasteFromState(destinationRow, destinationColumn); return}
            if (clipText === clipboardText) {pasteFromState(destinationRow, destinationColumn); return}
            if (!sameEntries && cut && clipText === clipboardText) { return}

            let row = Math.max(1, destinationRow)
            let column = Math.max(1, destinationColumn)
            let stringBuffer = ""

            function pushEntry(row, column, entry) {
                if (!newEntries[row]) {
                    newEntries[row] = []
                }
                newEntries[row][column] = entry
                stringBuffer = ""
            }

            if (clipboardText === clipText && cut && sameEntries) {
                for(let i=startRow; i<=endRow; i++) {
                    for(let j=startColumn; j<=endColumn; j++) {
                        if (newEntries[i]) newEntries[i][j] = ''
                    }
                }
            }

            for(let i=0; i < clipText.length; i++) {
                if (clipText.charCodeAt(i) === 9) {
                    if (stringBuffer.startsWith('"')) {
                        if (stringBuffer.endsWith('"') && stringBuffer.length > 1) {
                            stringBuffer = stringBuffer.substring(1, stringBuffer.length-1).replaceAll('""', '"');
                            pushEntry(row, column, stringBuffer)
                            column += 1
                        } else {
                            stringBuffer += clipText.charAt(i)
                        }
                    } else {
                        pushEntry(row, column, stringBuffer)
                        column += 1
                    }
                } else if (clipText.charCodeAt(i) === 13) {
                    if (stringBuffer.startsWith('"')) {
                        if (stringBuffer.endsWith('"') && stringBuffer.length > 1) {
                            stringBuffer = stringBuffer.substring(1, stringBuffer.length-1).replaceAll('""', '"');
                            pushEntry(row, column, stringBuffer)
                            row += 1
                            column = destinationColumn
                        } else {
                            stringBuffer += "\n"
                        }
                    } else {
                        pushEntry(row, column, stringBuffer)
                        row += 1
                        column = destinationColumn
                    }
                } else if (clipText.charCodeAt(i) != 10) {
                    stringBuffer += clipText.charAt(i)
                }

                if (i === clipText.length - 1) {
                    if (stringBuffer.startsWith('"') && stringBuffer.endsWith('"') && stringBuffer.length > 1) {
                        stringBuffer = stringBuffer.substring(1, stringBuffer.length-1).replaceAll('""', '"');
                    } 

                    pushEntry(row, column, stringBuffer)
                }
            }
            if (!cut || (sameEntries && clipboardText === clipText)) {
                setSelectedCells([{startRow: destinationRow, endRow: Math.min(row, file.rowCount), startColumn: destinationColumn, endColumn: Math.min(column, file.columnCount)}])
            }

            if (cut) {
                const promiseStart = Date.now()
                navigator.permissions.query({ name: "clipboard-write" }).then((result) => {
                    if (result.state === "granted" || result.state === "prompt") {
                        const promiseEnd = Date.now()

                        if (promiseEnd - promiseStart > 100) return

                        navigator.clipboard.writeText("").then(
                            () => {
                                /* clipboard successfully set */
                            },
                            () => {
                                /* clipboard write failed */
                                console.error("clipboard write failed !")
                            },
                        );
                    } else {
                        console.error(result)
                    }
                });
            }
        }

        function execute() {
            if (!cut && copyEntries && copyEntries.length === 1 && copyEntries[0].length === 1) {
                const lastSelection = getLastSelection()
                const [lastSelectionStartRow, lastSelectionEndRow] = ascendingOrder(lastSelection.startRow, lastSelection.endRow)
                const [lastSelectionStartColumn, lastSelectionEndColumn] = ascendingOrder(lastSelection.startColumn, lastSelection.endColumn)
    
                for(let i=0; i<=lastSelectionEndRow-lastSelectionStartRow; i++) {
                    for(let j=0; j<=lastSelectionEndColumn-lastSelectionStartColumn; j++) {
                        copyEntries = JSON.parse(JSON.stringify(copyData.entries))
                        const destinationRow = lastSelectionStartRow+i
                        const destinationColumn = lastSelectionStartColumn+j
    
                        if (copyEntries[0][0] != getCellValue(startRow, startColumn)) {
                            sameEntries = false
                        }
    
                        if (!cut) {
                            copyEntries[0][0] = incrementReferences(copyEntries[0][0], destinationRow-startRow, destinationColumn-startColumn)
                        }
    
                        paste(destinationRow, destinationColumn)
                    }
                }
            } else {
                const lastSelection = getLastSelection()
                let destinationRow
                let destinationColumn
                if (!lastSelection) {
                    destinationRow = selectedCell.row
                    destinationColumn = selectedCell.column
                } else {
                    const lastSelectionStartRow = ascendingOrder(lastSelection.startRow, lastSelection.endRow)[0]
                    const lastSelectionStartColumn = ascendingOrder(lastSelection.startColumn, lastSelection.endColumn)[0]
                    destinationRow = lastSelectionStartRow
                    destinationColumn = lastSelectionStartColumn
                }
    
                if (copyEntries) {
                    for(let i=0; i<copyEntries.length; i++) {
                        for (let j=0; j<copyEntries[i].length; j++) {
                            if (copyEntries[i][j] != getCellValue(i+startRow, j+startColumn)) {
                                sameEntries = false
                            }
        
                            if (!cut) {
                                copyEntries[i][j] = incrementReferences(copyEntries[i][j], destinationRow-startRow, destinationColumn-startColumn)
                            }
                        }
                    }
                }
    
                paste(destinationRow, destinationColumn)

                if (cut) {
                    for(let i=0; i<newEntries.length; i++) {
                        if (newEntries[i] === undefined) continue

                        for(let j=0; j<newEntries[i].length; j++) {
                            if (newEntries[i][j] === undefined) continue

                            let pos = 0
                            let quotes = false
                            let fixedRowReference = false
                            let fixedColumnReference = false
                            let rowReferenceStart = undefined
                            let columnReferenceStart = undefined
                            let columnStarted = false
                            
                            while (pos < newEntries[i][j].length) {
                                const char = newEntries[i][j][pos]
                                const nextChar = newEntries[i][j][pos+1]

                                if (!quotes && char === '"') {
                                    quotes = true;
                                } else if (quotes && char === '"') {
                                    if (nextChar && nextChar === '"') {
                                        pos += 2;
                                        continue;
                                    } else {
                                        quotes = false;
                                    }
                                }

                                if (!quotes) {
                                    /*console.log("---------------")
                                    console.log(pos)
                                    console.log(char)
                                    console.log(columnReferenceStart)
                                    console.log(rowReferenceStart)*/
                                    if ((char.match(/[a-z]/i) || char === "$") && (!columnReferenceStart || !columnStarted)) {
                                        if (columnReferenceStart === undefined) {
                                            columnReferenceStart = pos
                                        }
                                        if (char === "$") {
                                            fixedColumnReference = true
                                            pos += 1
                                            continue
                                        } else {
                                            columnStarted = true
                                        }
                                    } else if ((isDigit(char) || (char === "$")) && columnStarted && rowReferenceStart === undefined) {
                                        rowReferenceStart = pos
                                        if (char === "$") {
                                            fixedRowReference = true
                                            pos += 1
                                            continue
                                        }
                                    } else if (!(isDigit(char) || char.match(/[a-z]/i) || (char === "$"))) {
                                        if (rowReferenceStart !== undefined) {
                                            let column = newEntries[i][j].substring(columnReferenceStart, rowReferenceStart).replaceAll("$", "")
                                            let row = newEntries[i][j].substring(rowReferenceStart, pos).replaceAll("$", "")

                                            const initialLength = newEntries[i][j].substring(columnReferenceStart, pos).length
                                            column = excelColumnToNumber(column)
                                            row = parseInt(row)
                                            if (!isNaN(row) && !isNaN(column) && startRow <= row && row <= endRow && startColumn <= column && column <= endColumn) {
                                                row = destinationRow + (row-startRow)
                                                row = row.toString()
                                                column = destinationColumn + (column-startColumn)
                                                column = numberToExcelColumn(column)
                                                const newReference = (fixedColumnReference ? "$" : "") + column + (fixedRowReference ? "$" : "") + row
                                                newEntries[i][j] = newEntries[i][j].substring(0, columnReferenceStart) + newReference + newEntries[i][j].substring(pos)
                                                pos += newReference.length - initialLength - 1
                                            }
                                        }
                                        fixedRowReference = false
                                        fixedColumnReference = false
                                        rowReferenceStart = undefined 
                                        columnReferenceStart = undefined
                                        columnStarted = false
                                        pos +=1
                                        continue
                                    }

                                    if (rowReferenceStart != undefined && pos === newEntries[i][j].length-1) {
                                        let column = newEntries[i][j].substring(columnReferenceStart, rowReferenceStart).replaceAll("$", "")
                                        let row = newEntries[i][j].substring(rowReferenceStart, pos+1).replaceAll("$", "")

                                        column = excelColumnToNumber(column)
                                        row = parseInt(row)
                                        if (!isNaN(row) && !isNaN(column) && startRow <= row && row <= endRow && startColumn <= column && column <= endColumn) {
                                            row = destinationRow + (row-startRow)
                                            row = row.toString()
                                            column = destinationColumn + (column-startColumn)
                                            column = numberToExcelColumn(column)
                                            newEntries[i][j] = newEntries[i][j].substring(0, columnReferenceStart) + (fixedColumnReference ? "$" : "") + column + (fixedRowReference ? "$" : "") + row + newEntries[i][j].substring(pos+1)
                                        }
                                        break
                                    }
                                } else {
                                    fixedRowReference = false
                                    fixedColumnReference = false
                                    rowReferenceStart = undefined 
                                    columnReferenceStart = undefined
                                    columnStarted = false
                                }

                                pos += 1
                            }
                        }
                    }
                }
            }
    
            hideContextMenu()
    
            onChange({...file, inputs: newEntries})
        }

        if (!isFirefox()) {
            navigator.clipboard
            .readText()
            .then((clip) => {
                clipText = clip
                execute()
            }).catch((reason) => {
                execute()
            })
        } else {
            execute()
        }
    }
    function onContextMenuCopyValues() {
        if (!isFirefox()) {
            navigator.permissions.query({ name: "clipboard-write" }).then((result) => {
                if (result.state === "granted" || result.state === "prompt") {
                    const lastSelection = getLastSelection()

                    let clipboardText = ""
                    const [startRow, endRow] = ascendingOrder(lastSelection.startRow, lastSelection.endRow)
                    const [startColumn, endColumn] = ascendingOrder(lastSelection.startColumn, lastSelection.endColumn)
                    for(let i=startRow; i<=endRow; i++) {
                        for(let j=startColumn; j<=endColumn; j++) {
                            const cellValue = getComputedCellValue(i, j)
                            clipboardText += '"' + cellValue.replaceAll('"', '""') + '"'
                            if (j !== endColumn) {
                                clipboardText += String.fromCharCode(9)
                            }
                        }
                        if (i !== endRow) {
                            clipboardText += String.fromCharCode(13)
                        }
                    }

                    navigator.clipboard.writeText(clipboardText).then(
                        () => {
                            /* clipboard successfully set */
                            hideContextMenu()
                        },
                        () => {
                            /* clipboard write failed */
                            console.error("clipboard write failed !")
                        },
                    );
                } else {
                    console.error(result)
                    hideContextMenu()
                }
            });
        } else {
            hideContextMenu()
        }

        const lastSelection = getLastSelection()

        const copyEntries = []
        const [startRow, endRow] = ascendingOrder(lastSelection.startRow, lastSelection.endRow)
        const [startColumn, endColumn] = ascendingOrder(lastSelection.startColumn, lastSelection.endColumn)
        for(let i=startRow; i<=endRow; i++) {
            for(let j=startColumn; j<=endColumn; j++) {
                const cellValue = getComputedCellValue(i, j)
                copyEntries[i-startRow] = copyEntries[i-startRow] || []
                copyEntries[i-startRow][j-startColumn] = cellValue
            }
        }
        setCopyData({selection: lastSelection, entries: copyEntries})
    }

    let lastSelection = getLastSelection()
    let rowInsertCount = 1
    let columnInsertCount = 1
    let lastSelectionStartRow, lastSelectionEndRow
    let lastSelectionStartColumn, lastSelectionEndColumn

    if (lastSelection) {
        rowInsertCount = Math.min(file.rowCount, Math.abs(lastSelection.endRow-lastSelection.startRow)+1)
        columnInsertCount = Math.min(file.columnCount, Math.abs(lastSelection.endColumn-lastSelection.startColumn)+1)
        let result = ascendingOrder(lastSelection.startRow, lastSelection.endRow)
        lastSelectionStartRow = Math.max(1, result[0])
        lastSelectionEndRow = result[1]
        result = ascendingOrder(lastSelection.startColumn, lastSelection.endColumn)
        lastSelectionStartColumn = Math.max(1, result[0])
        lastSelectionEndColumn = result[1]
    }

    function insertRows(count) {
        const newEntries = [...file.inputs]

        for(let i=lastSelectionStartRow; i <= file.rowCount; i++) {
            for(let j=1; j <= file.columnCount; j++) {
                let pos = 0
                let quotes = false
                let fixedRowReference = false
                let rowReferenceStart = undefined
                let columnReferenceStart = undefined

                if (newEntries[i] === undefined || newEntries[i][j] === undefined) continue

                while (pos < newEntries[i][j]?.length) {
                    const char = newEntries[i][j][pos]
                    const nextChar = newEntries[i][j][pos+1]

                    if (!quotes && char === '"') {
                        quotes = true;
                    } else if (quotes && char === '"') {
                        if (nextChar && nextChar === '"') {
                            pos += 2;
                            continue;
                        } else {
                            quotes = false;
                        }
                    }

                    if (!quotes) {
                       if (char.match(/[a-z]/i) && !columnReferenceStart) {
                            columnReferenceStart = true
                       } else if (((char !== " " && isDigit(char)) || (char === "$")) && columnReferenceStart && !fixedRowReference) {
                            if (isDigit(char) && !rowReferenceStart) {
                                rowReferenceStart = pos
                            } else if (char === "$") {
                                fixedRowReference = true
                                pos += 1
                                continue
                            }
                       } else if (rowReferenceStart !== undefined) {
                            let row = newEntries[i][j].substring(rowReferenceStart, pos)
                            const initialLength = row.length
                            row = parseInt(row)
                            if (!isNaN(row) && row >= lastSelectionStartRow) {
                                row += count
                                row = row.toString()
                                newEntries[i][j] = newEntries[i][j].substring(0, rowReferenceStart) + row + newEntries[i][j].substring(pos)
                                pos += row.length - initialLength - 1
                            }
                            fixedRowReference = false
                            rowReferenceStart = undefined 
                            columnReferenceStart = undefined
                            pos +=1
                            continue
                       } else {
                            fixedRowReference = false
                            rowReferenceStart = undefined 
                            columnReferenceStart = undefined
                            pos +=1
                            continue
                       }

                       if (rowReferenceStart != undefined && pos === newEntries[i][j].length-1) {
                            let row = newEntries[i][j].substring(rowReferenceStart, pos+1)
                            row = parseInt(row)
                            if (!isNaN(row) && row >= lastSelectionStartRow) {
                                row += count
                                row = row.toString()
                                newEntries[i][j] = newEntries[i][j].substring(0, rowReferenceStart) + row + newEntries[i][j].substring(pos+1)
                            }
                            break
                       }
                    } else {
                        fixedRowReference = false
                        rowReferenceStart = undefined 
                        columnReferenceStart = undefined
                    }

                    pos += 1
                }
            }
        }

        if (count >= 0) {
            for(let i=0; i < rowInsertCount; i++) {
                newEntries.splice(lastSelectionStartRow, 0, [])
            }
        } else {
            for(let i=0; i < Math.abs(rowInsertCount); i++) {
                newEntries.splice(lastSelectionStartRow, 1)
            }
        }

        const newSelections = []
        for(let i=0; i<selectedCells.length; i++) {
            const selection = selectedCells[i]
            let [startRow, endRow] = ascendingOrder(selection.startRow, selection.endRow)
            startRow = Math.max(startRow, 1)
            endRow = Math.min(endRow, file.rowCount + count)

            if (startRow <= endRow) {
                newSelections.push({startRow: startRow, endRow: endRow, startColumn: selection.startColumn, endColumn: selection.endColumn})
            }
        }

        setCopyData({...copyData, hidden: true})
        setSelectedCell({row: Math.min(file.rowCount + count, Math.max(1, selectedCell.row)), column: selectedCell.column})
        setSelectedCells(newSelections)
        onChange({...file, inputs: newEntries, rowCount: file.rowCount + count})
        hideContextMenu()
    }

    function insertColumns(count) {
        const newEntries = [...file.inputs]

        for(let i=1; i <= file.rowCount; i++) {
            for(let j=lastSelectionStartColumn; j <= file.columnCount; j++) {
                let pos = 0
                let quotes = false
                let fixedReference = false
                let rowReferenceStart
                let columnReferenceStart

                if (newEntries[i] === undefined || newEntries[i][j] === undefined) continue

                while (pos < newEntries[i][j].length) {
                    const char = newEntries[i][j][pos]
                    const nextChar = newEntries[i][j][pos+1]

                    if (!quotes && char === '"') {
                        quotes = true;
                    } else if (quotes && char === '"') {
                        if (nextChar && nextChar === '"') {
                            pos += 2;
                            continue;
                        } else {
                            quotes = false;
                        }
                    }

                    if (!quotes) {
                        /*console.log("--------------------")
                        console.log("columnReferenceStart: " + columnReferenceStart?.toString())
                        console.log("rowReferenceStart: " + rowReferenceStart?.toString())
                        console.log("char: " + char?.toString())
                        console.log("fixedReference: " + fixedReference?.toString())*/

                       if (!columnReferenceStart && (char.match(/[a-z]/i) || char === "$")) {
                            columnReferenceStart = pos

                            if (char === "$") {
                                fixedReference = true
                                pos += 1
                                continue
                            }
                       } else if (((char !== " " && isDigit(char)) || char === "$") && columnReferenceStart) {
                            if ((isDigit(char) || char === "$") && !rowReferenceStart) {
                                rowReferenceStart = pos
                            }
                       } else if (rowReferenceStart !== undefined) {
                            let col = newEntries[i][j].substring(columnReferenceStart, rowReferenceStart)
                            const initialLength = col.length
                            col = excelColumnToNumber(col)
                            if (!isNaN(col) && col >= lastSelectionStartColumn) {
                                col += count
                                col = numberToExcelColumn(col)
                                newEntries[i][j] = newEntries[i][j].substring(0, columnReferenceStart) + col + newEntries[i][j].substring(rowReferenceStart)
                                pos += col.length - initialLength - 1
                            }
                            fixedReference = false
                            rowReferenceStart = undefined 
                            columnReferenceStart = undefined
                            pos +=1
                            continue
                       } else {
                            fixedReference = false
                            rowReferenceStart = undefined 
                            columnReferenceStart = undefined
                            pos +=1
                            continue
                       }

                       if (rowReferenceStart !== undefined && pos === newEntries[i][j].length-1) {
                            let col = newEntries[i][j].substring(columnReferenceStart, rowReferenceStart)
                            col = excelColumnToNumber(col)
                            if (!isNaN(col) && col >= lastSelectionStartColumn) {
                                col += count
                                col = numberToExcelColumn(col)
                                newEntries[i][j] = newEntries[i][j].substring(0, columnReferenceStart) + col + newEntries[i][j].substring(rowReferenceStart)
                            }
                            break
                       }
                    } else {
                        fixedReference = false
                        rowReferenceStart = undefined 
                        columnReferenceStart = undefined
                    }

                    pos += 1
                }
            }
        }

        if (count >= 0) {
            for(let i=0; i <= file.rowCount; i++) {
                if (!file.inputs[i]) continue

                for(let a=0; a < columnInsertCount; a++) {
                    newEntries[i].splice(lastSelectionStartColumn, 0, "")
                }
            }
        } else {
            for(let i=0; i <= file.rowCount; i++) {
                if (!file.inputs[i]) continue

                for(let a=0; a < Math.abs(columnInsertCount); a++) {
                    newEntries[i].splice(lastSelectionStartColumn, 1)
                }
            }
        }

        const newSelections = []
        for(let i=0; i<selectedCells.length; i++) {
            const selection = selectedCells[i]
            let [startColumn, endColumn] = ascendingOrder(selection.startColumn, selection.endColumn)
            startColumn = Math.max(startColumn, 1)
            endColumn = Math.min(endColumn, file.columnCount + count)

            if (startColumn <= endColumn) {
                newSelections.push({startRow: selection.startRow, endRow: selection.endRow, startColumn: startColumn, endColumn: endColumn})
            }
        }

        setCopyData({...copyData, hidden: true})
        setSelectedCell({row: selectedCell.row, column: Math.min(file.columnCount + count, Math.max(1, selectedCell.column))})
        setSelectedCells(newSelections)
        onChange({...file, inputs: newEntries, columnCount: file.columnCount + count})
        hideContextMenu()
    }

    function onContextMenuInsertRow() {
        insertRows(rowInsertCount)
        hideContextMenu()
    }
    function onContextMenuInsertColumn() {
        insertColumns(columnInsertCount)
        hideContextMenu()
    }
    function onContextMenuDeleteRow() {
        insertRows(-rowInsertCount)
        hideContextMenu()
    }
    function onContextMenuDeleteColumn() {
        insertColumns(-columnInsertCount)
        hideContextMenu()
    }


    function nextCell() {
        if (selectedCells.length === 1 && selectedCells[0].startColumn === selectedCells[0].endColumn && selectedCells[0].startRow === selectedCells[0].endRow) {
            const newRow = Math.min(selectedCell.row + 1, file.rowCount - 1)
            setSelectedCell({row: newRow, column: selectedCell.column})
            setSelectedCells([{startRow: newRow, startColumn: selectedCell.column, endRow: newRow, endColumn: selectedCell.column}])
            return
        }

        let currentSelectionObject = selectedCells[currentSelection]
        if (selectedCell.row + 1 <= currentSelectionObject.endRow) {
            setSelectedCell({...selectedCell, row: selectedCell.row + 1})
        } else {
            if (selectedCell.column + 1 <= currentSelectionObject.endColumn) {
                setSelectedCell({row: currentSelectionObject.startRow, column: selectedCell.column + 1})
            } else {
                if (currentSelection + 1 <= selectedCells.length - 1) {
                    setCurrentSelection(currentSelection + 1)
                    currentSelectionObject = selectedCells[currentSelection + 1]
                } else {
                    setCurrentSelection(0)
                    currentSelectionObject = selectedCells[0]
                }
                setSelectedCell({row: currentSelectionObject.startRow, column: currentSelectionObject.startColumn})
            }
        }
    }

    function onKeyDown(e) {
        if (e.key === "Enter") {
            if (selectedCells.length === 1 && selectedCells[0].startColumn === selectedCells[0].endColumn && selectedCells[0].startRow === selectedCells[0].endRow) {
                const cellRect = rowsRef.current.children[selectedCell.row].children[selectedCell.column].getBoundingClientRect()
                onDoubleClickCell(selectedCell.row, selectedCell.column, cellRect.left, cellRect.top)
            } else {
                nextCell()
            }
        } else if (e.key === "Delete" || e.key === "Backspace") {
            const newInputs = [...file.inputs]

            selectedCells.forEach((selection) => {
                for(let i=selection.startRow; i <= selection.endRow; i++) {
                    for(let j=selection.startColumn; j <= selection.endColumn; j++) {
                        if (file.inputs[i]) {
                            newInputs[i][j] = undefined

                            if (newInputs[i].length === 0) {
                                newInputs[i] = undefined
                            }
                        }
                    }
                }
            })
            onChange({...file, inputs: newInputs})
        } else if (e.key === "ArrowUp") {
            e.preventDefault()
            setSelectedCell({row: Math.max(selectedCell.row - 1, 1), column: selectedCell.column})
        } else if (e.key === "ArrowDown") {
            e.preventDefault()
            setSelectedCell({row: Math.min(selectedCell.row + 1, file.rowCount), column: selectedCell.column})
        } else if (e.key === "ArrowRight" || e.key === "Tab") {
            e.preventDefault()
            setSelectedCell({row: selectedCell.row, column: Math.min(selectedCell.column + 1, file.columnCount)})
        } else if (e.key === "ArrowLeft") {
            e.preventDefault()
            setSelectedCell({row: selectedCell.row, column: Math.max(selectedCell.column - 1, 1)})
        } else if (e.key.length === 1) {
            const cellRect = rowsRef.current.children[selectedCell.row].children[selectedCell.column].getBoundingClientRect()
            
            setCellEditorData({
                ...cellEditorData, 
                position: {x: cellRect.left, y: cellRect.top, row: selectedCell.row, column: selectedCell.column}, 
                value: ""
            })
        }
    }

    function onMouseUp(e) {
        if (selecting && !(extensionData && extensionData.extending)) {
            const newSelectedCells = [...selectedCells]
            const lastSelection = selectedCells[selectedCells.length - 1]
            if (!lastSelection) { return }

            const [lastSelectionStartRow, lastSelectionEndRow] = ascendingOrder(lastSelection.startRow, lastSelection.endRow)
            const [lastSelectionStartColumn, lastSelectionEndColumn] = ascendingOrder(lastSelection.startColumn, lastSelection.endColumn)

            let selectionRemoved = false

            selectedCells.forEach((selection, index) => {
                if (index === selectedCells.length - 1) { return }

                const [selectionStartRow, selectionEndRow] = ascendingOrder(selection.startRow, selection.endRow)
                const [selectionStartColumn, selectionEndColumn] = ascendingOrder(selection.startColumn, selection.endColumn)

                if (selectionStartColumn <= lastSelectionStartColumn && lastSelectionEndColumn <= selectionEndColumn
                    && selectionStartRow <= lastSelectionStartRow && lastSelectionEndRow <= selectionEndRow) {
                        if (!selectionRemoved) {
                            newSelectedCells.splice(newSelectedCells.length - 1, 1)
                        }
                        selectionRemoved = true

                        const additions = []
                        const selection = selectedCells[index]
                        const [selectionStartRow, selectionEndRow] = ascendingOrder(selection.startRow, selection.endRow)
                        const [selectionStartColumn, selectionEndColumn] = ascendingOrder(selection.startColumn, selection.endColumn)

                        if (selectionStartColumn === lastSelectionStartColumn && selectionEndColumn === lastSelectionEndColumn && selectionStartRow === lastSelectionStartRow && selectionEndRow === lastSelectionEndRow) {
                            newSelectedCells[index] = null
                            return
                        }

                        // First square
                        let startColumn = selectionStartColumn
                        let startRow = selectionStartRow
                        let endColumn = selectionEndColumn
                        let endRow = lastSelectionStartRow - 1

                        if (endColumn - startColumn >= 0 && endRow - startRow >= 0) {
                            additions.push({startColumn: startColumn, startRow: startRow, endColumn: endColumn, endRow: endRow})
                        }

                        // Second square
                        startColumn = selectionStartColumn
                        startRow = lastSelectionStartRow
                        endColumn = lastSelectionStartColumn - 1
                        endRow = lastSelectionEndRow

                        if (endColumn - startColumn >= 0 && endRow - startRow >= 0) {
                            additions.push({startColumn: startColumn, startRow: startRow, endColumn: endColumn, endRow: endRow})
                        }

                        // Third square
                        startColumn = lastSelectionEndColumn + 1
                        startRow = lastSelectionStartRow
                        endColumn = selectionEndColumn
                        endRow = lastSelectionEndRow

                        if (endColumn - startColumn >= 0 && endRow - startRow >= 0) {
                            additions.push({startColumn: startColumn, startRow: startRow, endColumn: endColumn, endRow: endRow})
                        }
                        
                        // Fourth square
                        startColumn = selectionStartColumn
                        startRow = lastSelectionEndRow + 1
                        endColumn = selectionEndColumn
                        endRow = selectionEndRow

                        if (endColumn - startColumn >= 0 && endRow - startRow >= 0) {
                            additions.push({startColumn: startColumn, startRow: startRow, endColumn: endColumn, endRow: endRow})
                        }
                        
                        newSelectedCells[index] = additions
                }
            })

            const formattedSelections = []
            newSelectedCells.forEach((value, i) => {
                if (value === null) {
                    return
                } else if (Array.isArray(value)) {
                    value.forEach((object) => {
                        formattedSelections.push(object)
                        return
                    })

                    if (i === newSelectedCells.length - 1) {
                        setCurrentSelection(formattedSelections.length - value.length)
                        setSelectedCell({row: formattedSelections[formattedSelections.length - value.length].startRow, column: formattedSelections[formattedSelections.length - value.length].startColumn})
                    }

                    return
                }

                formattedSelections.push(value)
            })
            setSelectedCells(formattedSelections)
            clearTimeout( timer );
        } else if (extensionData && extensionData.extending && extensionData.extension) {
            const extension = extensionData.extension
            const newEntries = [...file.inputs]

            let selectionStartRow, selectionEndRow, selectionStartColumn, selectionEndColumn
            const lastSelection = getLastSelection()
            if (lastSelection) {
                let result = ascendingOrder(lastSelection.startRow, lastSelection.endRow)
                selectionStartRow = result[0]
                selectionEndRow = result[1]
                result = ascendingOrder(lastSelection.startColumn, lastSelection.endColumn)
                selectionStartColumn = result[0]
                selectionEndColumn = result[1]
            } else {
                selectionStartRow = selectedCell.row
                selectionEndRow = selectedCell.row
                selectionStartColumn = selectedCell.column
                selectionEndColumn = selectedCell.column
            }

            if (extension) {
                if (extension.direction === "top" || extension.direction === "bottom") {
                    for(let j=selectionStartColumn; j<=selectionEndColumn; j++) {
                        let numberSequence = null
                        const dataPoints = []

                        for(let i=selectionStartRow; i<=selectionEndRow; i++) {
                            const expression = getCellValue(i, j)
                            const parsedExpression = parseInt(expression)

                            if (!isNaN(parsedExpression) || expression === "") {
                                if (numberSequence === null) {
                                    numberSequence = undefined
                                } else if (numberSequence === undefined) {
                                    numberSequence = true
                                }

                                if (!isNaN(parsedExpression)) {
                                    dataPoints.push({x: i, y: parsedExpression})
                                }
                            } else {
                                numberSequence = false
                                break
                            }
                        }

                        if (numberSequence === true && dataPoints.length >= 2) {
                            const result = linearTrend(dataPoints)

                            for(let v=extension.startRow; v<=extension.endRow; v++) {
                                const cellExpression = (result.slope * v + result.intercept).toString()
                                newEntries[v][j] = cellExpression
                            }
                        } else {
                            let i = extension.startRow

                            while(i <= extension.endRow) {
                                for(let v=selectionStartRow; v<=selectionEndRow; v++) {
                                    if (i > extension.endRow) break
                                    
                                    let cellExpression = getCellValue(v, j)

                                    cellExpression = incrementReferences(cellExpression, i-v, 0)
                                    newEntries[i][j] = cellExpression

                                    i++
                                }
                            }
                        }
                    }
                } else if (extension.direction === "left" || extension.direction === "right") {
                    for(let i=selectionStartRow; i<=selectionEndRow; i++) {
                        let numberSequence = null
                        const dataPoints = []

                        for(let j=selectionStartColumn; j<=selectionEndColumn; j++) {
                            const expression = getCellValue(i, j)
                            const parsedExpression = parseInt(expression)

                            if (!isNaN(parsedExpression) || expression === "") {
                                if (numberSequence === null) {
                                    numberSequence = undefined
                                } else if (numberSequence === undefined) {
                                    numberSequence = true
                                }

                                if (!isNaN(parsedExpression)) {
                                    dataPoints.push({x: j, y: parsedExpression})
                                }
                            } else {
                                numberSequence = false
                                break
                            }
                        }

                        if (numberSequence === true && dataPoints.length >= 2) {
                            const result = linearTrend(dataPoints)

                            for(let v=extension.startColumn; v<=extension.endColumn; v++) {
                                const cellExpression = (result.slope * v + result.intercept).toString()
                                newEntries[i][v] = cellExpression
                            }
                        } else {
                            let j = extension.startColumn

                            while(j <= extension.endColumn) {
                                for(let v=selectionStartColumn; v<=selectionEndColumn; v++) {
                                    if (j > extension.endColumn) break
                                    
                                    let cellExpression = getCellValue(i, v)
                                    cellExpression = incrementReferences(cellExpression, 0, j-v)
                                    newEntries[i][j] = cellExpression

                                    j++
                                }
                            }
                        }
                    }
                }
            }

            setSelectedCells([{
                startRow: Math.min(selectionStartRow, extension.startRow), 
                endRow: Math.max(selectionEndRow, extension.endRow), 
                startColumn: Math.min(selectionStartColumn, extension.startColumn),
                endColumn: Math.max(selectionEndColumn, extension.endColumn)
            }])
            onChange({...file, inputs: newEntries})
        }

        setExtensionData(null)
        setSelecting(false)
    }

    function onMouseDown(event) {
        if(contextMenuRef.current && !contextMenuRef.current.contains(event.target)){
            hideContextMenu()
        }
    }

    function getCellAlignment(column, row) {
        const computedValue = getComputedCellValue(column, row)
        return isPercentage(computedValue) || !isNaN(computedValue) ? "end" : "start"
    }

    function getColumnWidth(column) {
        return file.columnWidths[column] || file.defaultColumnWidth
    }

    function getRowHeight(row) {
        return file.rowHeights[row] || file.defaultRowHeight
    }

    function setCellValue(row, column, value) {
        const newInputs = [...file.inputs]
        if (!newInputs[row]) {
            newInputs[row] = []
        }
        newInputs[row][column] = value
        onChange({...file, inputs: newInputs})
    }

    function onCellEditorSubmit() {
        if (!cellEditorData.position) { return }
        setCellValue(cellEditorData.position.row, cellEditorData.position.column, cellEditorData.value)
    
        nextCell()
        rowsRef.current.focus()
    }

    const computedCells = computeCells(file.inputs, file.rowCount, file.columnCount)

    function getComputedCellValue(row, column) {
        if (computedCells[row] && computedCells[row][column]) {
            return computedCells[row][column]
        }

        return ""
    }

    function getCellValue(row, column) {
        if (file.inputs[row] && file.inputs[row][column]) {
            return file.inputs[row][column]
        }

        return ""
    }

    function onClickCell(button, row, column) {
        if (keyPressed.shift) return
        const lastSelection = getLastSelection()
        if (lastSelection) {
            const [lastStartColumn, lastEndColumn] = ascendingOrder(lastSelection.startColumn, lastSelection.endColumn)
            const [lastStartRow, lastEndRow] = ascendingOrder(lastSelection.startRow, lastSelection.endRow)
            if (button === 2 && lastSelection && ((lastStartRow <= row && row <= lastEndRow) && (lastStartColumn <= column && column <= lastEndColumn))) return
        }

        onCellEditorSubmit()
        let newSelectedCells
        if (keyPressed.ctrl) {
            newSelectedCells = [...selectedCells]
        } else {
            newSelectedCells = []
        }

        if (row === 0 && column === 0) {
            newSelectedCells = [{startRow: row, startColumn: column, endRow: file.rowCount, endColumn: file.columnCount}]
        } else if (row === 0) {
            let alreadySelected
            for(let i=0; i<selectedCells.length; i++) {
                const [startColumn, endColumn] = ascendingOrder(selectedCells[i].startColumn, selectedCells[i].endColumn)
 
                if (selectedCells[i].startRow === 0 && startColumn <= column && column <= endColumn) {
                    newSelectedCells.push({startColumn: -1, startRow: -1}) // Fake selection
                    alreadySelected = true
                    break
                }
            }

            if (!alreadySelected || !keyPressed.ctrl) {
                newSelectedCells[newSelectedCells.length] = {startRow: row, startColumn: column, endRow: file.rowCount, endColumn: column, headerColumnSelection: true}
            }
        } else if (column === 0) {
            let alreadySelected
            for(let i=0; i<selectedCells.length; i++) {
                const [startRow, endRow] = ascendingOrder(selectedCells[i].startRow, selectedCells[i].endRow)
 
                if (selectedCells[i].startColumn === 0 && startRow <= row && row <= endRow) {
                    newSelectedCells.push({startColumn: -1, startRow: -1}) // Fake selection
                    alreadySelected = true
                    break
                }
            }

            if (!alreadySelected || !keyPressed.ctrl) {
                newSelectedCells[newSelectedCells.length] = {startRow: row, startColumn: column, endRow: row, endColumn: file.columnCount, headerRowSelection: true}
            }
        } else {
            newSelectedCells[newSelectedCells.length] = {startRow: row, startColumn: column, endRow: row, endColumn: column}
            setSelectedCell({row: row, column: column})
        }

        setSelecting(true)
        setSelectedCells(newSelectedCells)
        setCurrentSelection(selectedCells.length - 1)
        setCellEditorData({...cellEditorData, position: null})
    }

    function selectCell(row, column) {
        if (widthResizerPosition != null) return
        if (heightResizerPosition != null) return
        if (cellEditorData && cellEditorData.position) return
        if (contextData && contextData.visible) return

        const selection = selectedCells[selectedCells.length - 1]
        if (!selection || selection.startColumn === 0 && selection.startRow === 0) return
        const newSelection = {...selection}

        if (selection.headerColumnSelection) {
            newSelection.endRow = file.rowCount
            newSelection.endColumn = column
        } else if (selection.headerRowSelection) {
            newSelection.endRow = row
            newSelection.endColumn = file.columnCount
        } else {
            newSelection.endRow = Math.max(row, 1)
            newSelection.endColumn = Math.max(column, 1)
        }
        
        const newSelectedCells = [...selectedCells]
        newSelectedCells[newSelectedCells.length - 1] = newSelection

        setSelectedCells(newSelectedCells)
    }

    function onSelectCell(row, column) {
        if (!(extensionData && extensionData.extending)) {
            selectCell(row, column)
        } else {
            let selectionStartRow, selectionEndRow, selectionStartColumn, selectionEndColumn
            const lastSelection = getLastSelection()
            if (lastSelection) {
                let result = ascendingOrder(lastSelection.startRow, lastSelection.endRow)
                selectionStartRow = result[0]
                selectionEndRow = result[1]
                result = ascendingOrder(lastSelection.startColumn, lastSelection.endColumn)
                selectionStartColumn = result[0]
                selectionEndColumn = result[1]
            } else {
                selectionStartRow = selectedCell.row
                selectionEndRow = selectedCell.row
                selectionStartColumn = selectedCell.column
                selectionEndColumn = selectedCell.column
            }

            const topDistance = selectionStartRow - row
            const rightDistance = column - selectionEndColumn
            const bottomDistance = row - selectionEndRow
            const leftDistance = selectionStartColumn - column

            let mostDistanceDirection = "top"
            let mostDistance = topDistance
            let extensionStartRow, extensionEndRow, extensionStartColumn, extensionEndColumn

            if (selectionStartRow <= row && row <= selectionEndRow && selectionStartColumn <= column && column <= selectionEndColumn) {
                setExtensionData({...extensionData, extension: null})
            } else {
                if (leftDistance > mostDistance) {
                    mostDistanceDirection = "left"
                    mostDistance = leftDistance
                }
                if (bottomDistance > mostDistance) {
                    mostDistanceDirection = "bottom"
                    mostDistance = bottomDistance
                }
                if (rightDistance > mostDistance) {
                    mostDistanceDirection = "right"
                    mostDistance = rightDistance
                }

                if (mostDistanceDirection === "top") {
                    extensionStartRow = selectionStartRow - topDistance
                    extensionEndRow = selectionStartRow - 1
                    extensionStartColumn = selectionStartColumn
                    extensionEndColumn = selectionEndColumn
                } else if (mostDistanceDirection === "right") {
                    extensionStartRow = selectionStartRow
                    extensionEndRow = selectionEndRow
                    extensionStartColumn = selectionEndColumn + 1
                    extensionEndColumn = selectionEndColumn + rightDistance
                } else if (mostDistanceDirection === "bottom") {
                    extensionStartRow = selectionEndRow + 1
                    extensionEndRow = selectionEndRow + bottomDistance
                    extensionStartColumn = selectionStartColumn
                    extensionEndColumn = selectionEndColumn
                } else if (mostDistanceDirection === "left") {
                    extensionStartRow = selectionStartRow
                    extensionEndRow = selectionEndRow
                    extensionStartColumn = selectionStartColumn - leftDistance
                    extensionEndColumn = selectionStartColumn - 1
                }

                setExtensionData({...extensionData, extension: {direction: mostDistanceDirection, startRow: extensionStartRow, endRow: extensionEndRow, startColumn: extensionStartColumn, endColumn: extensionEndColumn}})
            }
        }
    }

    function onDoubleClickCell(row, column, x, y) {
        setCellEditorData({
            ...cellEditorData, 
            position: {x: x, y: y, row: row, column: column}, 
            value: getCellValue(row, column)
        })
    }

    function onWidthResize(column, width) {
        const columnWidths = file.columnWidths
        columnWidths[column] = width
        onChange({...file, columnWidths: columnWidths})
    }
    
    function onHeightResize(row, height) {
        const rowHeights = file.rowHeights
        rowHeights[row] = height
        onChange({...file, rowHeights: rowHeights})
    }

    function onContextMenu(event) {
        const targetElement=mainRef.current
        if(targetElement && targetElement.contains(event.target)){
            event.preventDefault()
            if (event.target != contextMenuRef.current && !contextMenuRef.current.contains(event.target)) {
                if (cellEditorData && cellEditorData.position) return

                showContextMenu(event)
            }
        } else if(contextMenuRef.current && !contextMenuRef.current.contains(event.target)) {
            hideContextMenu()
        }
    }

    function startExtending() {
        setExtensionData({...extensionData, extending: true})
    }

    const selections = []

    if (selectedCells && !(selectedCells.length === 1 && selectedCells[0].startColumn === selectedCells[0].endColumn && selectedCells[0].startRow === selectedCells[0].endRow)) {
        selectionRects.forEach((selectionRect, i) => {
            selections.push(<Selection left={selectionRect.left} top={selectionRect.top} width={selectionRect.width} height={selectionRect.height} showExtender={i===selectionRects.length-1 && !cellEditorData.position} startExtending={startExtending} key={i}></Selection>)
        })
    }

    const rows = []

    for (let row = 0; row <= file.rowCount; row++) {
        const cells = []

        for (let column = 0; column <= file.columnCount; column++) {
            const firstRow = (row===0)
            const secondRow = (row===1)
            const lastRow = (row===file.rowCount)
            const firstColumn = (column===0)
            const secondColumn = (column===1)
            const beforeLastColumn = (column===file.columnCount)
            const lastColumn = (column===file.columnCount+1)

            const key=row.toString()+"-"+column.toString()

            // Insert block instead of editable cell at 1:1
            if (firstRow && firstColumn) {
                cells[column] = <EmptyBlock key={row.toString()+"-"+column.toString()} width={getColumnWidth(column)} height={getRowHeight(row)} onClickCell={onClickCell}></EmptyBlock>
                continue
            }

            // We must determine whether the cell occurs first or last in the row to determine whether or not it has left margin or right margin
            // The margin is useful because we want the resizer button to be fully contained in the cell div
            let positionInRow;
            if (firstColumn || (firstRow && secondColumn)) {
                positionInRow = "first"
            } else if (lastColumn || (firstRow && beforeLastColumn)) {
                positionInRow = "last"
            }

            // Same for top and bottom margin
            let positionInColumn;
            if (firstRow) {
                positionInColumn = "first"
            } else if (lastRow) {
                positionInColumn = "last"
            }

            // Generate the cell
            if (firstRow) {
                cells[column] = <HeaderRowCell 
                    width={getColumnWidth(column)} 
                    height={getRowHeight(row)} 
                    row={row}
                    column={column}
                    positionInRow={positionInRow}
                    positionInColumn={positionInColumn}
                    selected={selectedCell && selectedCell.row === row && selectedCell.column===column}
                    onClickCell={onClickCell}
                    onSelectCell={onSelectCell}
                    onWidthResize={onWidthResize} 
                    resizerPositionSetter={setWidthResizerPosition}
                    textAlign={getCellAlignment(row, column)}
                    key={key}
                >
                </HeaderRowCell>
            } else if (firstColumn) {
                cells[column] = <HeaderColumnCell 
                    width={getColumnWidth(column)} 
                    height={getRowHeight(row)} 
                    row={row}
                    column={column}
                    positionInRow={positionInRow}
                    positionInColumn={positionInColumn}
                    selected={selectedCell && selectedCell.row === row && selectedCell.column===column}
                    onClickCell={onClickCell}
                    onSelectCell={onSelectCell}
                    onHeightResize={onHeightResize} 
                    resizerPositionSetter={setHeightResizerPosition}
                    textAlign={getCellAlignment(row, column)}
                    key={key}
                >
                    {getComputedCellValue(row, column)} 
                </HeaderColumnCell>
            } else {   
                cells[column] = <Cell 
                    width={getColumnWidth(column)} 
                    height={getRowHeight(row)} 
                    row={row}
                    column={column}
                    positionInRow={positionInRow}
                    positionInColumn={positionInColumn}
                    selected={selectedCell && selectedCell.row === row && selectedCell.column===column}
                    onClickCell={onClickCell}
                    onSelectCell={onSelectCell}
                    onDoubleClickCell={onDoubleClickCell}
                    onWidthResize={onWidthResize} 
                    textAlign={getCellAlignment(row, column)}
                    showExtender={selections.length === 0 && selectedCell && selectedCell.row === row && selectedCell.column===column && !cellEditorData.position}
                    startExtending={startExtending}
                    key={key}
                >
                    {(!isNaN(getComputedCellValue(row, column)) && getComputedCellValue(row, column) != "") && row != 1 ? addCommas(Math.round(1000*getComputedCellValue(row, column))/1000) : getComputedCellValue(row, column)} 
                </Cell>
            }
        }
        
        rows[row] = <Row key={row} cells={cells}></Row>
    }

    let sameEntries = true
    if (copyData && !copyData.hidden) {
        const copyEntries = copyData.entries
        const selection = copyData.selection
        const startColumn = Math.max(1, ascendingOrder(selection.startColumn, selection.endColumn)[0])
        const startRow = Math.max(1, ascendingOrder(selection.startRow, selection.endRow)[0])

        for(let i=0; i<copyEntries.length; i++) {
            for (let j=0; j<copyEntries[i].length; j++) {
                if (copyEntries[i][j] != getCellValue(i+startRow, j+startColumn)) {
                    sameEntries = false
                    break
                }
            }
        }
    }

    let deleteRowsOptionText = "Delete row"
    if (lastSelectionStartRow !== lastSelectionEndRow) {
        deleteRowsOptionText += "s " + lastSelectionStartRow.toString() + "-" + lastSelectionEndRow.toString()
    }

    let deleteColumnsOptionText = "Delete column"
    if (lastSelectionStartColumn !== lastSelectionEndColumn) {
        deleteColumnsOptionText += "s " + numberToExcelColumn(lastSelectionStartColumn) + "-" + numberToExcelColumn(lastSelectionEndColumn)
    }

    return <div ref={scrollRef} style={{maxWidth: "96vw", maxHeight: "90vh", width: width, height: height, overflow: "auto"}}>
        <div ref={mainRef} style={{width: "fit-content", paddingRight: "50px", paddingBottom: "50px", cursor: "default"}}><div style={{position: "relative", width: "fit-content", cursor: contextData?.visible ? "default" : "cell"}} onContextMenu={onContextMenu}>
            {/* Display cell input if a cell is double clicked*/}
            {cellEditorData.position && <CellEditor 
                spreadsheetRef={mainRef}
                position={cellEditorData.position} 
                width={getColumnWidth(cellEditorData.position.column)} height={getRowHeight(cellEditorData.position.row)} 
                value={cellEditorData.value}
                onChange={(value) => setCellEditorData({...cellEditorData, value: value})}
                onSubmit={onCellEditorSubmit}
                onExit={() => setCellEditorData({...cellEditorData, position: null})}>
            </CellEditor>}

            {/* Display bar at the right of cell if resizing width*/}
            {widthResizerPosition && <div
                style={{
                    position: "absolute",
                    width: 1,
                    height: "100%",
                    left: widthResizerPosition,
                    backgroundColor: "black",
                    zIndex: 2,
                }}
            >     
            </div>}

            {/* Display bar at the bottom of cell if resizing height*/}
            {heightResizerPosition && <div
                style={{
                    position: "absolute",
                    width: "100%",
                    height: 1,
                    top: heightResizerPosition,
                    backgroundColor: "black",
                    zIndex: 2,
                }}
            >     
            </div>}

            {/* Display the rows containing the cells*/}
            <div ref={rowsRef} onKeyDown={onKeyDown} tabIndex={0} style={{outline: "none", width: "fit-content"}}>
                {rows}
            </div>

            <div ref={selectionsRef}>
                {selections}
                {copyData && copyData.selectionRect && !copyData.hidden && sameEntries && <CopySelection left={copyData.selectionRect.left} top={copyData.selectionRect.top} width={copyData.selectionRect.width} height={copyData.selectionRect.height}></CopySelection>}
                {(extensionData && extensionData.rect && extensionData.extension) && <Extension direction={extensionData.extension.direction} left={extensionData.rect.left} top={extensionData.rect.top} width={extensionData.rect.width} height={extensionData.rect.height}></Extension>}
            </div>

            <ContextMenu visible={contextData.visible} posX={contextData.posX} posY={contextData.posY} refProp={contextMenuRef}
                options={[
                    [
                        true ? <ContextMenuOption onClick={onContextMenuCut} key="cut">Cut</ContextMenuOption> : null,
                        true ? <ContextMenuOption onClick={onContextMenuCopy} key="copy">Copy</ContextMenuOption> : null,
                        true ? <ContextMenuOption onClick={onContextMenuPaste} key="paste">Paste</ContextMenuOption> : null,
                    ],
                    [
                        true ? <ContextMenuOption onClick={onContextMenuCopyValues} key="copyValues">Copy values only</ContextMenuOption> : null
                    ],
                    [
                        true ? <ContextMenuOption onClick={onContextMenuInsertRow} key="insertRow">{'Insert '+(rowInsertCount === 1 ? 'a' : rowInsertCount)+' '+(rowInsertCount > 1 ? 'rows' : 'row')+' above'}</ContextMenuOption> : null,
                        true ? <ContextMenuOption onClick={onContextMenuInsertColumn} key="insertColumn">{'Insert '+(columnInsertCount === 1 ? 'a' : columnInsertCount)+' '+(columnInsertCount > 1 ? 'columns' : 'column')+' on the left'}</ContextMenuOption> : null,
                    ],
                    [
                        true ? <ContextMenuOption onClick={onContextMenuDeleteRow} key="deleteRow">{deleteRowsOptionText}</ContextMenuOption> : null,
                        true ? <ContextMenuOption onClick={onContextMenuDeleteColumn} key="deleteColumn">{deleteColumnsOptionText}</ContextMenuOption> : null,
                    ]
                ]}
            >
            </ContextMenu>
        </div></div>
    </div>
}