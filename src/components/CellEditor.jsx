import React, { useEffect, useRef } from 'react';

function setCaretToEnd(el, pos) {
    var range = document.createRange()
    var sel = window.getSelection()

    let lastTextNode
    for(let i=el.childNodes.length-1; i>=0; i--) {
        if (el.childNodes[i].nodeType === Node.TEXT_NODE) {
            lastTextNode = el.childNodes[i]
            break
        }
    }
    if (!lastTextNode || !lastTextNode.data) return
    
    range.setStart(lastTextNode, lastTextNode.data.length)
    range.collapse(true)
    
    sel.removeAllRanges()
    sel.addRange(range)
}

export default function CellEditor({width, height, position, value, onChange, onSubmit, onExit, spreadsheetRef}) {
    const myRef = useRef(null)

    useEffect(() => {
        myRef.current.innerText = value
        setCaretToEnd(myRef.current)
        myRef.current.focus();
        autoGrow()
      }, [position]);

    function onCellChange(e) {
        autoGrow()
        onChange(e.target.innerText)
    }

    function onKeyDown(e) {
        if (e.key === "Escape") {
            onExit()
        } else if (e.key === "Enter") {
            onSubmit()
            onExit()
        }
    }

    // Block pasting hmtl in the Cell Editor
    function onPaste(e) {
        e.preventDefault();
        const text = e.clipboardData.getData("text/plain");
        document.execCommand("insertText", false, text);
    }

    function autoGrow(event) {
        const textarea = myRef.current;
        
        // Calculate the width based on the content
        const contentWidth = getTextWidth(textarea.innerText, getComputedStyle(textarea).font);

        //const newWidth = Math.min(window.innerWidth - textarea.getBoundingClientRect().width + width, contentWidth ); // Limit to the right screen border
        //textarea.style.width = newWidth + 'px';
        
        textarea.style.width = 'fit-content';
        textarea.style.width = Math.min(contentWidth, window.innerWidth - textarea.getBoundingClientRect().left-4, spreadsheetRef.current.getBoundingClientRect().left + spreadsheetRef.current.clientWidth - textarea.getBoundingClientRect().left-4-8) + 'px';
        
        // Calculate the height based on the content
        textarea.style.height = 'fit-content';
        textarea.style.overflow = "hidden"
        let contentHeight = textarea.scrollHeight - 1
        const windowLimitedHeight =  window.innerHeight - textarea.getBoundingClientRect().top-4
        const spreadsheetLimitedHeight = spreadsheetRef.current.getBoundingClientRect().top + spreadsheetRef.current.clientHeight - textarea.getBoundingClientRect().top-4
        if (contentHeight > windowLimitedHeight) {
            contentHeight = windowLimitedHeight
            textarea.style.overflow = "auto"
        }
        if (contentHeight > spreadsheetLimitedHeight) {
            contentHeight = spreadsheetLimitedHeight
            textarea.style.overflow = "auto"
        }

        textarea.style.height = contentHeight + 'px';
    }
  
    function getTextWidth(text, font) {
        const canvas = document.createElement('canvas');
        const context = canvas.getContext('2d');
        context.font = font;
        const width = context.measureText(text).width;
      
        // Clean up the canvas element
        canvas.remove();
      
        return width;
    }

    return <div className='resize-container' style={{left: (position && position.x) || 0,
        top: (position && position.y) || 0, position: "fixed", zIndex:3,}}>
        <div suppressContentEditableWarning={true} contentEditable onPaste={onPaste} className="resize-input" onKeyDown={onKeyDown} onInput={onCellChange} hidden={position === null} ref={myRef} type="text" style={{
            display: "block",
            type: "text",
            overflow: "auto",
            minWidth: width - 4,
            minHeight: height,
            padding: 0,
            outline: "none",
            borderWidth: 2,
            borderColor: "blue",
            borderStyle: "solid",
            borderRadius: 0,
            resize: "none",
            textAlign: "left",
            backgroundColor: "white",
            verticalAlign: "top",
            wordWrap : "break-word" // Don't change to overflowWrap or otherwise setCaretToEnd will not work
        }} ></div>
    </div>
}