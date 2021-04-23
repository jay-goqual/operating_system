import React, { useState, useEffect } from 'react';
import { TransitionGroup, CSSTransition } from 'react-transition-group';
import { Spinner, Table, Button, Col, Row, ListGroup, Form, ToggleButton, ButtonGroup } from 'react-bootstrap';
import Select from 'react-select';
import DaumPostcode from 'react-daum-postcode';

import FormInput from './FormInput.tsx';

import server from '../../utils/server';
import { getProducts } from '../../../server/sheets';

const { serverFunctions } = server;

const SearchOrder = () => {
    const [data, setData] = useState([]);
    const [check, setCheck] = useState([]);
    const [searching, setSearching] = useState(false);
    const [typecheck, setTypecheck] = useState([]);
    const [memo, setMemo] = useState([]);
    const [send, setSend] = useState([]);
    const [isAddress, setIsAddress] = useState();
    const [isZoneCode, setIsZoneCode] = useState();
    const [extraAddress, setExtraAddress] = useState();
    const [isPostOpen, setIsPostOpen] = useState(false);
    const [products, setProducts] = useState();
    const [test, setTest] = useState();

    const findOrder = async input => {
        try {
            setSearching(true);
            const response = await serverFunctions.findOrder(input);
            setData(response);
            setSearching(false);
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

    const cs_type = [{name: '단순반품', value: 1}, {name: '보상반품', value: 2}, {name: '교환', value: 3}, {name: '재작업', value: 4}, {name: '재발송', value: 5}];

    const handleSubmit = () => {
        // setMemo('');
        // setTypecheck(0);
        // setCheck([]);
        // // setData([]);
        // setSend([]);
        alert(test.value);
    };

    const handleSingleCheck = (checked, id) => {
        if (checked) {
            setCheck([...check, id]);
        } else {
            setCheck(check.filter((el) => el !== id));
        }
    };

    const addProduct = value => {
        alert(value.value)
        let temp = send;
        temp[0].value = value.value;
        temp[0].lable = value.label;
        setSend(temp);
    }

    const addSend = () => {
        setSend([...send, {value: '', label: '', num: 0}]);
    }

    const deleteSend = () => {
        setSend(send.slice(0, -1));
    }

    const handleRadio = (e) => {
        setTypecheck(e.currentTarget.value);
        if (check.length > 0) {
            setIsAddress(data[check[0]].customer_address);
            setIsZoneCode(data[check[0]].customer_zipcode);
        }
    }

    const handleComplete = (data) => {
        let fullAddress = data.address;
        let extraAddress = "";
    
        if (data.addressType === "R") {
            if (data.bname !== "") {
                extraAddress += data.bname;
            }
            if (data.buildingName !== "") {
                extraAddress +=
                extraAddress !== "" ? `, ${data.buildingName}` : data.buildingName;
            }
            fullAddress += extraAddress !== "" ? ` (${extraAddress})` : "";
        }
        setIsZoneCode(data.zonecode);
        setIsAddress(fullAddress);
        setIsPostOpen(false);
    };

    const handlePost = () => {
        setIsPostOpen(!isPostOpen);
    }

    const postCodeStyle = {
        width: "100%",
        height: '500px',
        padding: "7px",
    };
    

    useEffect(() => {
        const fetch = async() => {
            try {
                const result = await serverFunctions.getProducts();
                setProducts(result);
            } catch (error) {
                alert(error);
            }
        }
        fetch();
        // setProducts(serverFunctions.getProducts());
    }, []);


    return (
        <div style={{ padding: '3px', fontSize: 12, width: 1490}}>
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
            <FormInput findOrder={findOrder}/>
            <div>
                {searching ? 
                    <div style={{display: 'flex', alignItems: 'center', justifyContent: 'center'}}>
                        <Spinner animation="border" role="status" />
                    </div> :
                    <Table striped bordered hover size="sm">
                        <thead>
                            <tr>
                                <th className='check'>
                                    <div className='data_div'>
                                        <Form>
                                            <Form.Check type="checkbox" checked="true" disabled />
                                        </Form>
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
                                            <Form>
                                                <Form.Check type="checkbox" onChange={(e) => handleSingleCheck(e.target.checked, key)} checked={check.includes(key) ? true : false} />
                                            </Form>
                                        </td>
                                        {Object.keys(list_header).map((k) => (
                                            <td className={k}>
                                                <div className='data_div'>
                                                    {data[key][k]}
                                                </div>
                                            </td>
                                        ))}
                                    </tr>
                                ))}
                        </tbody>
                    </Table>
                }

                <hr />
                <div>
                    <ButtonGroup toggle>
                        {cs_type.map((k, i) => (
                            <ToggleButton size='sm' type='radio' key={i} variant="outline-secondary" value={k.value} checked={typecheck == k.value} onChange={handleRadio}>
                                {k.name}
                            </ToggleButton>
                        ))}
                    </ButtonGroup>
                </div>
                <div style={{marginTop: 15}}>
                    <Form.Group>
                        <Form.Label className='title'>메모</Form.Label>
                        <Form.Control as="textarea" rows={2} value={memo} onChange={(e) => setMemo(e.currentTarget.value)} />
                    </Form.Group>
                </div>

                {(typecheck == 3 || typecheck == 5) &&
                    <div>
                        <hr />
                        <div className='title'>
                            상품 발송
                        </div>
                        <div style={{marginTop: 10}}>
                            <Form style={{marginLeft: 100, marginRight: 100}}>
                                <Row sty>
                                    <Col>
                                        <Form.Label size='sm'>수령인명</Form.Label>
                                        <Form.Control size='sm' type='text' style={{width: 200}} />
                                    </Col>
                                    <Col>
                                        <Form.Label size='sm'>연락처</Form.Label>
                                        <Form.Control size='sm' type='text' style={{width: 200}} />
                                    </Col>
                                    <Col /><Col /><Col /><Col />
                                </Row>
                                <br />
                                <Form.Label size='sm'>주소</Form.Label>
                                <Row>
                                    <Col>
                                        <Form.Control size='sm' type='text' value={isAddress} disabled />
                                    </Col>
                                    <Col>
                                        <Form.Control size='sm' type='text' value={extraAddress} onChange={(e) => setExtraAddress(e.currentTarget.value)} />
                                    </Col>
                                    <Col>
                                        <Form.Control size='sm' type='text' value={isZoneCode} disabled style={{width: 100}} />
                                    </Col>
                                </Row>
                                <br />
                                <Button size='sm' variant='outline-secondary' onClick={handlePost}>검색</Button>
                                {isPostOpen && <DaumPostcode style={postCodeStyle} onComplete={handleComplete} />}
                                <br />
                                <br />
                                <div style={{display: 'flex', justifyContent: 'space-between'}}>
                                    <Form.Label size='sm'>제품</Form.Label>
                                    <div style={{display: 'flex'}}>
                                        <Button size='sm' variant="outline-secondary" onClick={addSend}>
                                            추가
                                        </Button>
                                        <div style={{width: 5}} />
                                        <Button size='sm' variant="outline-secondary" onClick={deleteSend}>
                                            제거
                                        </Button>
                                    </div>
                                </div>
                                {send.map((s, i) => (
                                    <div style={{width: 1000, display: 'flex', marginTop: 10}} key={i}>
                                        <div style={{width: 600, marginRight: 20}}>
                                            <Select autoFocus options={products} value={s} onChange={addProduct} />
                                        </div>
                                        <div style={{width: 60}}>
                                            <Form.Control type="text" />
                                        </div>
                                    </div>
                                ))}
                            </Form>
                        </div>
                        {/* {Object.keys(send).map((i) => (
                            <div key={i}>
                                test
                            </div>
                        ))} */}
                    </div>
                }

                <hr />
                <div>
                    {/* <Button variant="primary" type="submit" onClick={handleSubmit} disabled={!(check.length > 0 && typecheck > 0)}> */}
                    <Button variant="primary" size='sm' type="submit" onClick={handleSubmit}>
                        제출
                    </Button>
                </div>

            </div>
        </div>
    );
};

export default SearchOrder;
