export function CopySelection({left, top, width, height}) {
    return <div style={{
        position: "absolute",
        left: left - 2,
        top: top - 2,
        width: width + 1,
        height: height +1,
        pointerEvents: "none",
        borderStyle: "dashed",
        borderColor: "rgb(26, 115, 232)",
        borderWidth: 1,
    }}>
    </div>
}

export function Selection({left, top, width, height, showExtender, startExtending}) {
    function onMouseDownExtender(event) {
        startExtending()
    }

    return <div style={{
        position: "absolute",
        left: left - 1,
        top: top - 1,
        width: width - 1,
        height: height - 1,
        backgroundColor: "rgba(0, 98, 255, 0.1)",
        borderStyle: "solid",
        borderColor: "rgb(26, 115, 232)",
        borderWidth: 1,
        pointerEvents: "none"
    }}>
    {showExtender && <div onMouseDown={onMouseDownExtender} style={{cursor: "crosshair", userSelect: false, pointerEvents: "auto", position: "absolute", bottom: 0, right: 0, transform: "translate(50%, 50%)", zIndex: 1, width: "8px", height: "8px", borderRadius: "50%", backgroundColor: "blue", borderWidth: 1, borderStyle: "solid", borderColor: "white"}}></div>}
    </div>
}

export function Extension({direction, left, top, width, height}) {
    return <div style={{
        position: "absolute",
        left: left - 1,
        top: top - 1,
        width: width -1,
        height: height - 1,
        pointerEvents: "none",
        borderStyle: "dashed",
        borderColor: "black",
        borderWidth: "1px",
        borderTopStyle: (direction === "bottom" ? "none" : "dashed"),
        borderRightStyle: (direction === "left" ? "none" : "dashed"),
        borderBottomStyle: (direction === "top" ? "none" : "dashed"),
        borderLeftStyle: (direction === "right" ? "none" : "dashed")
    }}>
    </div>
}