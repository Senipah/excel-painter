import styled from "styled-components";

const Container = styled.div`
  display: flex;
  flex-direction: column;
  padding: 0.5rem;
  margin: 0.5rem;
  align-items: center;
  width: 100%;
  max-width: 540px;

  /* Medium devices (landscape tablets, 768px and up) */
  @media only screen and (min-width: 768px) {
    max-width: 720px;
  }

  /* Large devices (laptops/desktops, 992px and up) */
  @media only screen and (min-width: 992px) {
    max-width: 960px;
  }

  /* Extra large devices (large laptops and desktops, 1200px and up) */
  @media only screen and (min-width: 1200px) {
    max-width: 1200px;
  }
`;

export default Container;
