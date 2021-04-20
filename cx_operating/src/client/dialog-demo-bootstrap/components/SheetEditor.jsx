import React, { useState, useEffect } from 'react';
import { TransitionGroup, CSSTransition } from 'react-transition-group';
import { Table, Button, ListGroup } from 'react-bootstrap';
import FormInput from './FormInput.tsx';

// This is a wrapper for google.script.run that lets us use promises.
import server from '../../utils/server';

const { serverFunctions } = server;

const SheetEditor = () => {
  const [data, setData] = useState([]);

  const findOrder = async input => {
    try {
        const response = await serverFunctions.findOrder(input);
        setData(response);
    } catch (error) {
        alert(error);
    }
  };

  /* useEffect(() => {
    serverFunctions
      .getSheetsData()
      .then(setNames)
      .catch(alert);
  }, []);

  const deleteSheet = sheetIndex => {
    serverFunctions
      .deleteSheet(sheetIndex)
      .then(setNames)
      .catch(alert);
  };

  const setActiveSheet = sheetName => {
    serverFunctions
      .setActiveSheet(sheetName)
      .then(setNames)
      .catch(alert);
  };

  const submitNewSheet = async newSheetName => {
    try {
      const response = await serverFunctions.addSheet(newSheetName);
      setNames(response);
    } catch (error) {
      // eslint-disable-next-line no-alert
      alert(error);
    }
  }; */

  return (
    <div style={{ padding: '3px', overflowX: 'hidden' }}>
      <p>
        <b>☀️ Bootstrap demo! ☀️</b>
      </p>
      <p>
        This is a sample app that uses the <code>react-bootstrap</code> library
        to help us build a simple React app. Enter a name for a new sheet, hit
        enter and the new sheet will be created. Click the red{' '}
        <span className="text-danger">&times;</span> next to the sheet name to
        delete it.
      </p>
      <FormInput findOrder={findOrder} />
      {/* <ListGroup>
        <TransitionGroup className="sheet-list">
          {names.length > 0 &&
            names.map(name => (
              <CSSTransition
                classNames="sheetNames"
                timeout={500}
                key={name.name}
              >
                <ListGroup.Item
                  className="d-flex"
                  key={`${name.index}-${name.name}`}
                >
                  <Button
                    className="border-0"
                    variant="outline-danger"
                    size="sm"
                    onClick={() => deleteSheet(name.index)}
                  >
                    &times;
                  </Button>
                  <Button
                    className="border-0 mx-2"
                    variant={name.isActive ? 'success' : 'outline-success'}
                    onClick={() => setActiveSheet(name.name)}
                  >
                    {name.name}
                  </Button>
                </ListGroup.Item>
              </CSSTransition>
            ))}
        </TransitionGroup>
      </ListGroup> */}
        {/* <ListGroup>
            <TransitionGroup className="orderList">
                {data.length > 0 &&
                    data.map(order => (
                    <CSSTransition
                        classNames="sheetNames"
                        timeout={500}
                        key={order.order_uid}
                    >
                        <ListGroup.Item
                            className="d-flex"
                            key={order.order_uid}
                        >
                            {order}
                        </ListGroup.Item>
                    </CSSTransition>
                ))}
          </TransitionGroup>
      </ListGroup> */}
        <div class="table-container">
        <Table striped bordered hover responsive size="sm">
            <tbody>
                {data.length > 0 &&
                    Object.keys(data).map((k) => (
                        <tr key={`${k}`}>
                            {Object.keys(data[k]).map((kk) => (
                                <td key={`${k}-${kk}`}>
                                    {data[k][kk]}
                                </td>
                            ))}
                        </tr>
                    ))}
            </tbody>
        </Table>
        </div>
    </div>
  );
};

export default SheetEditor;
