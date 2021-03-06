import React, { useState, ChangeEvent, FormEvent } from 'react';
import { Form, Button, Col, Row } from 'react-bootstrap';

interface FormInputProps {
  findOrder: (
    input: string
  ) => {

  };
  /* ) => {
    name: string;
    index: number;
    isActive: boolean;
  }; */
}

const FormInput = ({ findOrder }: FormInputProps) => {
  const [input, setInput] = useState('');

  const handleChange = (event: ChangeEvent<HTMLInputElement>) =>
    setInput(event.target.value);

  const handleSubmit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (input.length === 0) return;
    findOrder(input);
    setInput('');
  };

  return (
    <Form onSubmit={handleSubmit}>
      <Form.Group controlId="formNewSheet">
        <Row>
          <Col xs={10}>
            <Form.Control
              type="text"
              placeholder="성함"
              size='sm'
              value={input}
              onChange={handleChange}
            />
          </Col>
          <Col xs={2}>
            <Button variant="primary" type="submit" size='sm'>
              Submit
            </Button>
          </Col>
        </Row>
      </Form.Group>
    </Form>
  );
};

export default FormInput;
