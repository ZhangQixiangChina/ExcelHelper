Object.prototype.hasValue = function (name) {
    return !!this[name];
};

Object.prototype.keys = function () {
    return Object.keys(this);
};

Object.prototype.delete = function (key) {
    delete this[key];
};

Array.prototype.remove = function (index) {
    this.splice(index, 1);
};