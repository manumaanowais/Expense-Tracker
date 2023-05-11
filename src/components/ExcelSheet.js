import { useState, useEffect } from "react";
import * as XLSX from 'xlsx';

function ExcelSheet() {

  const [data, setData] = useState([]);
  useEffect(() => {
    const fetchData = async () => {
      const url = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vTuEAsp1wI-I5KGC5SFuY0xGtC4_aqlYkZzLNxqu0Q8eGWnfAbbJ86zkxosm51IePz3dNSIXgVUJPGB/pubhtml?gid=1588499464&single=true'; // replace with your Excel file URL
      const response = await fetch(url);
      const arrayBuffer = await response.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const sheetData = XLSX.utils.sheet_to_json(worksheet);
      setData(sheetData);
      //   console.warn(sheetData[0])
      let totalExpenses = sheetData.map(i => i.Amount)
      const calculateTotal = (total, num) => {
        return total + num
      }
      document.getElementById('loading').innerHTML = 'Total Expenses : '
      document.getElementById('totalExpense').innerHTML = totalExpenses.reduce(calculateTotal)
    };
    fetchData()
  }, [])

  const amountStyle = {
    color: 'green',
    marginLeft: '300px',
    marginTop: '-45px'
  }

  const handleDelete = (row, index) => {
    let updatedTotalExpenses = document.getElementById('totalExpense').innerHTML - row.Amount;
    const updatedData = [...data];
    updatedData.splice(index, 1);
    setData(updatedData);
    document.getElementById('totalExpense').innerHTML = updatedTotalExpenses
    console.warn('Row with id ', index + 1, ' deleted')

  }

  let idCounter = 1;
  return (
    <div>
      <button className="btn btn-primary"><a className='expenseForm' href="https://docs.google.com/forms/d/e/1FAIpQLSeUQflwPyqgZvuIL5qtZ_myMgEbenbiTjUvpP70-SsScfMcYg/viewform">Add Expense</a></button>
      <table id='expenseTable' className="table table-striped table-bordered table-sm" cellSpacing="0" width="100%">
        <thead>
          <tr>
            <th scope="col">Id</th>
            <th scope="col">Date</th>
            <th scope="col">Month</th>
            <th scope="col">Salary</th>
            <th scope="col">Expense</th>
            <th scope="col">Amount</th>
            <th scope="col">Actions</th>
          </tr>
        </thead>
        <tbody>
          {data.map((row, index) => (
            <tr key={index}>
              <th scope="row">{idCounter++}</th>
              <td>{new Intl.DateTimeFormat('en-US', { year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' }).format(row.timestamp)}</td>
              <td>{row.Month}</td>
              <td>{row.Salary}</td>
              <td>{row.Expense}</td>
              <td>{row.Amount}</td>
              <td>
                <button type="button" className="btn btn-info"><a className="editForm" href={row.Editlinks}>Edit</a></button>&nbsp;
                <button type="button" className="btn btn-danger" onClick={() => handleDelete(row, index)}>Delete</button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>

      <h1 id='loading'>Loading....</h1>
      <h2 id='totalExpense' style={amountStyle}>.</h2>
    </div>
  );
}

export default ExcelSheet;
