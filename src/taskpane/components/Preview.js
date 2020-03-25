import * as React from "react";
import styled from "styled-components";
import { PrimaryButton, ButtonType, Spinner, SpinnerType } from "office-ui-fabric-react";

const Card = styled.div`
  display: flex;
  flex-direction: column;
  justify-content: center;
  align-items: center;
  width: 100%;
  border-radius: 0.25rem;
  border: 1px solid rgba(0, 0, 0, 0.125);
  background-color: paleblue;
  margin: 1rem;
  & > * {
    margin: 0.5rem;
  }
`;

const CardBody = styled.div`
  display: flex;
  flex-direction: column;
  padding: 1.25rem;
`;

const Image = styled.img`
  height: 100%;
  width: 100%;
  display: block;
`;

const Wrapper = styled.div`
  display: flex;
  flex-direction: column;
  justify-content: center;
  align-items: center;
  width: 100%;
  & > * {
    margin: 1rem;
  }

  @media only screen and (min-width: 600px) {
    flex-direction: row;
    align-items: flex-end;
  }
`;

const Title = styled.h2`
  word-wrap: break-word;
`;

const Preview = props => {
  const { image, handleClick, busy } = props;
  return (
    <Card>
      <CardBody>
        <Title>{image.displayName}</Title>
        <Wrapper>
          {busy ? (
            <Spinner type={SpinnerType.large} label="Working" />
          ) : (
            <PrimaryButton
              className="ms-welcome__action"
              buttonType={ButtonType.hero}
              disabled={busy}
              iconProps={{ iconName: "ChevronRight" }}
              onClick={handleClick}
            >
              Run
            </PrimaryButton>
          )}
        </Wrapper>
      </CardBody>
      <Image src={image.src} alt={image.fileName} />
    </Card>
  );
};

export default Preview;
