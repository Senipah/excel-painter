import * as React from "react";
// import { TextField, Button } from "office-ui-fabric-react";
import styled from "styled-components";

const ControlWrapper = styled.div`
  display: flex;
  flex-direction: column;
  justify-content: center;
  align-items: center;
  width: 100%;
  & > * {
    margin: 0.5rem;
  }

  & > input {
    flex-grow: 2;
  }

  @media only screen and (min-width: 600px) {
    flex-direction: row;
    align-items: flex-end;
  }
`;

const FilePicker = props => {
  return (
    <ControlWrapper className="ms-TextField">
      <label htmlFor="image" className="ms-Label">
        Choose a picture
      </label>
      <input
        className="ms-TextField-field"
        type="file"
        id="image"
        name="image"
        accept="image/*"
        onChange={props.handleChange}
      />
    </ControlWrapper>
  );
};

export default FilePicker;
