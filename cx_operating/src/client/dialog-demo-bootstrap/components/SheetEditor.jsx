import React, { useState, useEffect } from 'react';
import { TransitionGroup, CSSTransition } from 'react-transition-group';
import { Table, Button, ListGroup, Form, ToggleButton } from 'react-bootstrap';
import FormInput from './FormInput.tsx';

// This is a wrapper for google.script.run that lets us use promises.
import server from '../../utils/server';

const { serverFunctions } = server;

const SheetEditor = () => {
  const [data, setData] = useState([]);
  const [check, setCheck] = useState([]);

  const findOrder = async input => {
    try {
        const response = await serverFunctions.findOrder(input);
        /* let c = {};
        response.map((v, i) => {
            c[i] = false;
        }); */
        setData(response);
        // setCheck(c);
    } catch (error) {
        alert(error);
    }
  };

  const list_header = {
      'date_receipt': '접수일',
      'seller_name': '판매처',
      'order_id': '주문번호',
      'order_uid': '상품주문번호',
      'order_name': '주문자',
      'order_phone': '연락처',
      'customer_name': '수취인',
      'customer_phone': '연락처',
      'customer_address': '주소',
      'customer_zipcode': '우편번호',
      'product_code': '상품코드',
      'product_name': '상품명',
      'product_num': '수량',
      'order_option': '옵션'
  };

  const handleSubmit = () => {
      Object.keys(check).map((v, i) => {
          alert(check[v]);
      })
  };

  const handleChecked = key => {
    //   alert(key);
      let temp = check;
      temp[key] = !temp[key];
      setCheck(temp);
    //   alert(check[key]);
  };

  const handleSingleCheck = (checked, id) => {
    if (checked) {
      setCheck([...check, id]);
    } else {
      // 체크 해제
      setCheck(check.filter((el) => el !== id));
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
        <Table striped bordered hover size="sm">
            <thead>
                <tr>
                    <th className='check'>
                        <div className='data_div'>
                            <Form.Check type="checkbox" checked="true" disabled />
                        </div>
                    </th>
                    {Object.keys(list_header).map((k) => (
                        <th className={k}>
                            <div className='data_div'>
                                {list_header[k]}
                            </div>
                        </th>
                    ))}
                </tr>
            </thead>
            <tbody>
                {data.length > 0 &&
                    Object.keys(data).map((key) => (
                        <tr>
                            <td className='check'>
                                <Form.Check type="checkbox" onChange={(e) => handleSingleCheck(e.target.checked, key)} checked={check.includes(key) ? true : false} />
                                {/* <ToggleButton type="checkbox" checked={check[key]} onClick={handleChecked(key)} /> */}
                            </td>
                            {Object.keys(list_header).map((k) => (
                                <td className={k}>
                                    <div className='data_div'>
                                        {data[key][k]}
                                    </div>
                                </td>
                            ))}
                            {/* {Object.keys(data[k]).map((kk) => (
                                <td key={`${k}-${kk}`}>
                                    {data[k][kk]}
                                </td>
                            ))} */}
                        </tr>
                    ))}
            </tbody>
        </Table>
        <Button variant="primary" type="submit" onClick={handleSubmit}>
            제출
        </Button>
    </div>
  );
};

export default SheetEditor;
