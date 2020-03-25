import * as React from "react";
import styled from "styled-components";

const FlexSection = styled.section`
  display: flex
  justify-content: center;
  align-items: center:
  flex-direction: column;
  text-align: center;
`;

export default class Header extends React.Component {
  render() {
    const { title, logo, message } = this.props;

    return (
      <FlexSection>
        <img width="90" height="90" src={logo} alt={title} title={title} />
        <h1 className="ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary">{message}</h1>
      </FlexSection>
    );
  }
}
