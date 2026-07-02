export default class Fields {
  constructor(headerResponse = {}, detailResponse = {}) {
    this.data = {
      ...(headerResponse.data ?? {}),
      Details: detailResponse.data ?? [],
    };

    this.labels = {
      ...(headerResponse.labels ?? {}),
      ...(detailResponse.labels ?? {}),
    };
  }

  get(key) {
    return this.data[key];
    
  }

  getLabel(key) {
    return this.labels[key];
  }
}