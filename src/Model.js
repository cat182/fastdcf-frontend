import logo from './logo.svg';
import './Model.css';
import Spreadsheet from './components/Spreadsheet';
import { ToastContainer } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';
import { useState } from "react"
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome'
import { faHouse } from '@fortawesome/free-solid-svg-icons'
import { MathJax, MathJaxContext } from 'better-react-mathjax';
import { useLoaderData } from 'react-router-dom';
import { generateEntries } from './functions/financialParser'

export async function loader({ params }) {
  const response = await fetch("http://localhost:3001/entities/"+params.entity);

  if (response.status === 404) {
    return { error: 404 }
  }

  const entity = await response.json()
  return { entity };
}

function App() {
  const { entity, error } = useLoaderData();

  const [entries, rowCount, columnCount] = generateEntries(entity.incomeStatements, entity.cashFlowStatements, entity.balanceSheets, entity.maximumDivider)

  const [spreadsheetFile, setSpreadsheetFile] = useState({
    rowCount: rowCount,
    columnCount: columnCount,
    columnWidths: [undefined, 190],
    inputs: JSON.parse(JSON.stringify(entries))
  })

  if (error === 404) {
    return <div>
      Entity not found
    </div>
  }

  return (
    <MathJaxContext>
      <div className="App" style={{paddingLeft: "30px", paddingRight: "30px", textAlign: "left"}}>
        <header style={{position: "relative", paddingTop: "16px", paddingBottom: "16px", height: 30}}>
          <FontAwesomeIcon className='fa-2xl' icon={faHouse} style={{position: "absolute", left: "0", top: "50%", transform: "translate(0%, -50%)"}} />
          <h1 className='fullTitle' style={{textAlign: "center", margin: 0, position: "absolute", left: "50%", top: "50%", transform: "translate(-50%, -50%)"}}>{"Entity: "+entity.name}</h1>
          <h1 className='shortTitle' style={{textAlign: "center", margin: 0, position: "absolute", left: "50%", top: "50%", transform: "translate(-50%, -50%)"}}>{"Entity: "+entity.name}</h1>
          <div className='headerButtonContainer' style={{position: "absolute", right: "0", top: "50%", transform: "translate(0%, -50%)"}}>
            <button style={{width: "70px"}}>Log In</button>
            <button className='signupButton' style={{width: "70px"}}>Sign Up</button>
          </div>
        </header>
        <div style={{display: "flex", justifyContent: "center", marginBottom: 5}}>
          <Spreadsheet file={spreadsheetFile} onChange={setSpreadsheetFile} height="100vh"></Spreadsheet>
        </div>
        <div>
          <div>
            <label>Terminal Growth Rate (%): </label>
            <input type="number" style={{width: 125}} step={0.01}></input>
          </div>
          <hr></hr>
          <div>
            <label>Effective Tax Rate (%): </label>
            <input type="number" style={{width: 125}} readOnly disabled></input>
            <i> We use the average Tax Rate computed from the Income Statements</i>
          </div>

          <br></br>

          <div style={{marginBottom: 5, display: "flex"}}>
            <label>Use custom Effective Tax Rate: </label>
            <input type="checkbox"/>
          </div>
          <div hidden style={{marginTop: 5}}>
            <label>Custom Effective Tax Rate (%): </label>
            <input type="number" style={{width: 125}} step={0.01}></input>
          </div>
          <hr></hr>
          <div style={{marginBottom: 5}}>
            <label>Levered Beta: </label>
            <input type="number" style={{width: 125}} readOnly disabled></input>
            <i> We use Yahoo Finance's Betas</i>
          </div>

          <div style={{marginBottom: 5}}>
            <label>Avg. Debt/Equity Ratio: </label>
            <input type="number" style={{width: 125}} readOnly disabled></input>
            <i> The average Debt/Equity Ratio during the period of regression of the Beta computed from the Balance Sheets</i>
          </div>

          <div style={{marginBottom: 5}}>
            <label>Unlevered Beta: </label>
            <input type="number" style={{width: 125}} readOnly disabled></input>
            <i> Levered Beta / (1 + (1 - Effective Tax Rate) * Avg. Debt/Equity Ratio)</i>
          </div>

          <div style={{marginBottom: 5}}>
            <label>Unlevered Beta without Cash: </label>
            <input type="number" style={{width: 125}} readOnly disabled></input>
            <i> Unlevered Beta / (1 - Avg. Cash Balance as % of Firm Value)</i>
          </div>

          <div style={{marginBottom: 5}}>
            <label>Current Debt/Equity Ratio: </label>
            <input type="number" style={{width: 125}} readOnly disabled></input>
            <i> We get it from the last filing</i>
          </div>

          <div style={{marginBottom: 5}}>
            <label>Levered Beta without Cash: </label>
            <input type="number" style={{width: 125}} readOnly disabled></input>
            <i> Unlevered Beta without Cash * (1 + (1 - Effective Tax Rate) * (Current Debt/Equity Ratio))</i>
          </div>

          <div style={{marginBottom: 5}}>
            <label>Risk-free Rate (%): </label>
            <input type="number" style={{width: 125}} readOnly disabled></input>
            <i> We use "CBOE Interest Rate 10 Year T No" value (10 Year U.S. Treasury Note Yield) from Yahoo Finance</i>
          </div>

          <div style={{marginBottom: 5}}>
            <label>Equity Risk Premium (%): </label>
            <input type="number" style={{width: 125}} readOnly disabled></input>
            <i> We use Aswath Damodaran's implied equity risk premium</i>
          </div>

          <div>
            <label>CAPM Cost of Equity (%): </label>
            <input type="number" style={{width: 125}} readOnly disabled></input>
            <i> Risk-free rate + Levered Beta without Cash * Equity Risk Premium</i>
          </div>

          <br></br>

          <div style={{marginBottom: 5, display: "flex"}}>
            <label>Use custom Cost of Equity: </label>
            <input type="checkbox"/>
          </div>
          <div hidden style={{marginTop: 5}}>
            <label>Custom Cost of Equity (%): </label>
            <input type="number" style={{width: 125}} step={0.01}></input>
          </div>
          <hr></hr>
          <div>
            <label>Cost of Debt (%): </label>
            <input type="number" style={{width: 125}} readOnly disabled></input>
            <i> We aproximate it by last filing (Interest Expense / Book Total Debt) for practical purposes</i>
          </div>

          <br></br>

          <div style={{display: "flex"}}>
            <label>Use custom Cost of Debt: </label>
            <input type="checkbox"/>
          </div>
          <hr></hr>
          <div style={{marginBottom: 5}}>
            <label>Market Cap (in thousands): </label>
            <input type="number" style={{width: 125}} readOnly disabled></input>
            <i> We get it from Yahoo Finance</i>
          </div>

          <div style={{marginBottom: 5}}>
            <label>Market value of Debt (in thousands): </label>
            <input type="number" style={{width: 125}} step={1} readOnly disabled></input>
            <i> We aproximate it by Book Total Debt from last filing for practical purposes</i>
          </div>

          <div style={{marginBottom: 5}}>
            <label>Weighted Average Cost of Capital (%): </label>
            <input type="number" style={{width: 125}} readOnly disabled></input>
            <i> ((Market Cap / (Market Cap + Market Value of Debt)) * Cost of Equity) + ((Market Value of Debt / (Market Cap + Market Value of Debt)) * Cost of Debt * Effective Tax Rate)</i>
          </div>

          <br></br>

          <div style={{marginBottom: 5, display: "flex"}}>
            <label>Use custom WACC: </label>
            <input type="checkbox"/>
          </div>
          <div hidden style={{marginTop: 5}}>
            <label>Custom WACC (%): </label>
            <input type="number" style={{width: 125}} step={0.01}></input>
          </div>

          <hr></hr>
          <div style={{display: "flex", flexDirection: "row", alignItems: "center"}}>
            <label>Terminal Value (in thousands):</label>
            <input type="number" style={{width: 125, height: "fit-content", marginLeft: "3px"}} readOnly disabled></input>
            <MathJax style={{width: "fit-content", marginLeft: "3px"}}>{ "\\[\\frac{\\text{FCFF}_{\\text{Last Forecasted Year}} \\times (1 + \\text{TGR})}{\\text{WACC} - \\text{TGR}}\\]" }</MathJax>
          </div>

          <div style={{display: "flex", flexDirection: "row", alignItems: "center"}}>
            <label>Enterprise Value (in thousands):</label>
            <input type="number" style={{width: 125, height: "fit-content", marginLeft: "3px"}} readOnly disabled></input>
            <MathJax style={{width: "fit-content", marginLeft: "3px"}}>{ "\\[\\sum_{t=1}^\\text{Number of Periods} \\frac{\\text{FCFF}}{(1 + \\text{WACC})^t} + \\frac{\\text{Terminal Value}}{(1 + \\text{WACC})^\\text{Number of Periods}}\\]" }</MathJax>
          </div>

          <div>
            <label>Net Cash (in thousands): </label>
            <input type="number" style={{width: 125}} readOnly disabled></input>
            <i> We get it from the last filing: (Cash and Cash equivalents + Short-Term Investments) - Total Debt</i>
          </div>

          <br></br>

          <div style={{marginBottom: 5, display: "flex"}}>
            <label>Use custom Net Cash: </label>
            <input type="checkbox"/>
          </div>
          <div hidden style={{marginTop: 5}}>
            <label>Custom Net Cash (in thousands): </label>
            <input type="number" style={{width: 125}} step={1}></input>
          </div>

          <hr></hr>

          <div style={{marginBottom: 5}}>
            <label>Shares Outstanding (in nominal terms): </label>
            <input type="number" style={{width: 125}} readOnly disabled></input>
            <i> We get it from the Yahoo Finance</i>
          </div>

          <br></br>

          <div style={{marginBottom: 5, display: "flex"}}>
            <label>Use custom Shares Outstanding: </label>
            <input type="checkbox"/>
          </div>
          <div hidden style={{marginTop: 5}}>
            <label>Custom Shares Outstanding (in nominal terms): </label>
            <input type="number" style={{width: 125}} step={1}></input>
          </div>

          <hr></hr>

          <div>
            <label>Present Value of a Share: </label>
            <input type="number" style={{width: 125}} readOnly disabled></input>
            <i> (Entreprise Value + Net Cash) / Shares Outstanding</i>
          </div>
        </div>

        <br></br>

        Mettre interest income and interest expense a zero et mettre le lien de damodaran dans la cellule
        <ToastContainer />
      </div>
    </MathJaxContext>
  );
}

export default App;
