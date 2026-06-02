globalThis.foundry = {
  utils: {
    deepClone: (o) => JSON.parse(JSON.stringify(o)),
  },
};
