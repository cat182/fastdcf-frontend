import React, { useState, useRef, useEffect } from "react"
import { numberToExcelColumn } from "../functions/util";

var mouseDown = false;
document.onmousedown = function() { 
  mouseDown = true;
}
document.onmouseup = function() {
  mouseDown = false;
}

export function Row({cells}) {
    return <div className="react-row">
        {cells}
    </div>
}

export function EmptyBlock({width, height, onClickCell}) {
    function onCellMouseDown(event) {
        onClickCell(event.button, 0, 0)
    }

    return <div
        onMouseDown={onCellMouseDown}
        style={{
            minWidth: width, 
            minHeight: height,
            maxHeight: height,
            maxWidth: width,
            backgroundColor: "rgb(240, 241, 240)",
        }}
        className='react-cell-tres'
    > 
    </div>
}

export function HeaderRowCell({row, column, width, height, positionInRow, positionInColumn, resizerPositionSetter, onWidthResize, textAlign, onClickCell, onSelectCell}) {
    const [resizerHovered, setResizerHovered] = useState(false)
    const [resizingWidth, setResizingWidth] = useState(false)

    const myRef = useRef(null)

    useEffect(() => {
        document.addEventListener(
            "mousemove",
            onMouseMove
        )

        return () => {
            document.removeEventListener("mousemove", onMouseMove)
        }
    }, [resizingWidth])

    function onMouseDown(event) {
        event.preventDefault()
        setResizingWidth(true)
        
        document.addEventListener(
            "mouseup",
            (e) => {
                setResizingWidth(false)
                const cellRect = myRef.current.getBoundingClientRect()
                onWidthResize(column, Math.max(e.clientX - cellRect.left, 10))
                resizerPositionSetter(null)
            },
            { once: true }
          );
    }

    function onMouseMove(e) {
        if (resizingWidth) {
            const cellRect = myRef.current.getBoundingClientRect()
            resizerPositionSetter(Math.max(e.clientX - cellRect.left, 10) + cellRect.left - myRef.current.parentNode.parentNode.parentNode.getBoundingClientRect().left)
        }
    }

    function onResizerEnter() {
        setResizerHovered(true)
    }

    function onResizerLeave() {
        setResizerHovered(false)
    }

    function onCellMouseDown(event) {
        onClickCell(event.button, row, column)
    }

    function onCellEnter() {
        if (mouseDown) {
            onSelectCell(row, column)
        }
    }

    return <div
            ref={myRef}
            onMouseEnter={onCellEnter}
            onMouseDown={onCellMouseDown}
            style={{
                minWidth: width,
                minHeight: height,
                maxWidth: width,
                maxHeight: height,
                marginRight: positionInRow !== "last" ? 8 : 0,
                marginLeft: positionInRow !== "first" ? -8 : 0  ,
                marginBottom: positionInColumn !== "last" ? 8 : 0,
                marginTop: positionInColumn !== "first" ? -8 : 0,
                position: "relative",
                userSelect: "none",
                backgroundColor: "rgb(240, 241, 240)"
            }}
            className="react-cell-quatro"
        >
            {<div
                style={{
                    width: 11,
                    right: -6,
                    height: 11,
                    top: height / 2 - 11 /2, 
                    margin:0,
                    borderColor: "black",
                    backgroundColor: resizerHovered ? "black" : null,
                    position: "absolute",
                    zIndex: 1,
                }}
                onMouseEnter={onResizerEnter}
                onMouseLeave={onResizerLeave}
                onMouseDown={onMouseDown}
            >
            </div>}
            <div style={{overflow: "hidden", height: height, textAlign: "center"}}>
                <span>{numberToExcelColumn(column)}</span>
            </div>
        </div>
}

export function HeaderColumnCell({children, row, column, width, height, positionInRow, positionInColumn, resizerPositionSetter, onHeightResize, textAlign, onSelectCell, onClickCell}) {
    const [resizerHovered, setResizerHovered] = useState(false)
    const [resizingHeight, setResizingHeight] = useState(false)

    const myRef = useRef(null)

    useEffect(() => {
        document.addEventListener(
            "mousemove",
            onMouseMove
        )

        return () => {
            document.removeEventListener("mousemove", onMouseMove)
        }
    }, [resizingHeight])

    function onMouseDown(event) {
        event.preventDefault()
        setResizingHeight(true)
        
        document.addEventListener(
            "mouseup",
            (e) => {
                setResizingHeight(false)
                const cellRect = myRef.current.getBoundingClientRect()
                onHeightResize(row, Math.max(e.clientY - cellRect.top, 10))
                resizerPositionSetter(null)
            },
            { once: true }
          );
    }

    function onMouseMove(e) {
        if (resizingHeight) {
            const cellRect = myRef.current.getBoundingClientRect()
            resizerPositionSetter(Math.max(e.clientY - cellRect.top, 10) + cellRect.top - myRef.current.parentNode.parentNode.parentNode.getBoundingClientRect().top )
        }
    }

    function onResizerEnter() {
        setResizerHovered(true)
    }

    function onResizerLeave() {
        setResizerHovered(false)
    }

    function onCellMouseDown(event) {
        onClickCell(event.button, row, column)
    }

    function onCellEnter() {
        if (mouseDown) {
            onSelectCell(row, column)
        }
    }

    return <div
        ref={myRef}
        onMouseEnter={onCellEnter}
        onMouseDown={onCellMouseDown}
        style={{
            minWidth: width,
            minHeight: height,
            maxWidth: width,
            maxHeight: height,
            marginRight: positionInRow !== "last" ? 8 : 0,
            marginLeft: positionInRow !== "first" ? -8 : 0,
            marginBottom: positionInColumn !== "last" ? 8 : 0,
            marginTop: positionInColumn !== "first" ? -8 : 0,
            position: "relative",
            userSelect: "none",
            backgroundColor: "rgb(240, 241, 240)",
        }}
        className="react-cell-bis"
    >
        {<div
            style={{
                width: 11,
                bottom: -6,
                left: width / 2 - 11/2,
                height: 11,
                margin:0,
                borderColor: "black",
                backgroundColor: resizerHovered ? "black" : null,
                position: "absolute",
                zIndex: 1,
            }}
            onMouseEnter={onResizerEnter}
            onMouseLeave={onResizerLeave}
            onMouseDown={onMouseDown}
        >
        </div>}
        <div style={{overflow: "hidden", height: height, textAlign: "center"}}>
            <span>{row}</span>
        </div>
    </div>
}

export function Cell({row, column, children, width, height, onClickCell, onSelectCell, onDoubleClickCell, positionInRow, positionInColumn, selected, textAlign, showExtender, startExtending}) {
    const myRef = useRef(null)

    function onCellMouseDown(event) {
        onClickCell(event.button, row, column)
    }

    function internalOnDoubleClickCell(event) {
        const cellRect = myRef.current.getBoundingClientRect()
        onDoubleClickCell(row, column, cellRect.left, cellRect.top)
    }

    function onMouseDownExtender(event) {
        if (event.button !== 0) return
        startExtending()
    }

    function onCellEnter() {
        if (mouseDown) {
            onSelectCell(row, column)
        }
    }

    return <div
            ref={myRef}
            onMouseDown={onCellMouseDown}
            onDoubleClick={internalOnDoubleClickCell}
            onMouseEnter={onCellEnter}
            style={{
                minWidth: width,
                minHeight: height,
                maxWidth: width,
                maxHeight: height,
                marginRight: positionInRow !== "last" ? 8 : 0,
                marginLeft: positionInRow !== "first" ? -8 : 0,
                marginBottom: positionInColumn !== "last" ? 8 : 0,
                marginTop: positionInColumn !== "first" ? -8 : 0,
                position: "relative",
                userSelect: "none",
                overflowWrap: "break-word"
            }}
            className="react-cell"
        >
        {selected && <div style={{position: "absolute", borderStyle: "solid", marginLeft: 0, marginTop: 0, width: width-4, height: height-4, borderWidth: 2, borderColor: "blue"}}></div>}
        {showExtender && <div onMouseDown={onMouseDownExtender} style={{cursor: "crosshair", position: "absolute", bottom: 0, right: 0, transform: "translate(50%, 50%)", zIndex: 1, width: "8px", height: "8px", borderRadius: "50%", backgroundColor: "blue", borderWidth: 1, borderStyle: "solid", borderColor: "white"}}></div>}
        <div style={{overflow: "hidden", width: "100%", height: "100%"}}>
            <div style={{textOverflow:"clip", backgroundColor: "white", height: 20, textAlign: textAlign}}>
                <span style={{whiteSpace: "break-spaces"}}>{children}</span>
            </div>
        </div>
    </div>
}