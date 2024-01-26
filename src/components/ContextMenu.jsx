import { useState } from 'react';

export function ContextMenuOption({children, onClick}) {
    const [hovered, setHovered] = useState(null)

    return <div onClick={onClick} onMouseEnter={() => setHovered(true)} onMouseLeave={() => setHovered(false)} style={{cursor: "pointer", userSelect: "none", width: "100%", height: "25px", display: "flex", alignItems: "center", justifyContent: "center", backgroundColor: hovered ? "rgb(241, 242, 244)" : "white"}}>
        <span style={{margin: "auto"}}>{children}</span>
    </div>
}

export function ContextMenu({visible, posX, posY, options, refProp}) {
    const elements = []
    
    let hrIndex = 0

    options.forEach((optionGroup, i) => {
        optionGroup.forEach((option) => {
            elements.push(option)
        })
        if (optionGroup.length > 0 && i < options.length - 1) {
            elements.push(<hr key={hrIndex} />)
            hrIndex++;
        }
    });

    return <div ref={refProp} style={{zIndex:2, borderRadius: "5px", visibility: visible ? "visible" : "hidden", width: "200px", paddingTop: "10px", paddingBottom: "10px", position: "fixed",  backgroundColor: "white", left: posX, top: posY, boxShadow: "0 3px 10px rgb(0 0 0 / 0.2)"}}>
        {elements}
    </div>
}