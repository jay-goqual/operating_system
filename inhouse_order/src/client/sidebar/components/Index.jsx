import React, { useState, useEffect } from 'react';
import { TransitionGroup, CSSTransition } from 'react-transition-group';
import { Button, ListGroup } from 'react-bootstrap';
// import FormInput from './FormInput.tsx';

// This is a wrapper for google.script.run that lets us use promises.
import server from '../../utils/server';

const { serverFunctions } = server;

const Index = () => {
    const pushOrder = () => {
        serverFunctions
          .pushOrder()
          .catch(alert);
      };
  
  return (
    <div style={{ padding: '3px', overflowX: 'hidden' }}>
      <p>
        내부발주를 위한 스프레드시트입니다. <br /><br />
        <b>원활한 사용을 위해 다른 팀에서 사용 중일 경우에는, 사용이 완료된 이후 작성을 시작해주세요.</b> <br /><br />
        사용 중 궁금하신 점은, <code>커머스팀 신재유 매니저</code> 에게 문의 부탁드리겠습니다. <br />
      </p>
      <Button variant="primary" type="submit" onClick={() => pushOrder()}> 제출하기 </Button>
      <p className="text-link">
          <br />
          <br />
          <br />
          <a className="text-link" href = "https://search.naver.com/search.naver?where=nexearch&ie=utf8&X_CSA=address_search&query=%EC%9A%B0%ED%8E%B8%EB%B2%88%ED%98%B8" target="_blank" rel="noopener noreferrer"> ** 네이버 우편번호 찾기 </a>
          <br />
          <br />
          <p>** 첫 사용시에는 상단 [고퀄] 메뉴 내의 [제출하기]를 한 번 실행해주세요</p>
      </p>
    </div>
  );
};

export default Index;
