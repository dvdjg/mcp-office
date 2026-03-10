class MockWinaxObject {
  constructor(..._args: unknown[]) {}

  __release() {}
}

const winax = {
  Object: MockWinaxObject,
};

export default winax;
