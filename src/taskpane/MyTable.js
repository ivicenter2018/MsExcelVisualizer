// JavaScript source code

export default class MyTable {

    name;
    formulas;
    size;

    constructor(name) {
        this.name = name;
        this.formulas = new Map();
        this.size = 0;
        
    }

    newFormula(id, newFormula) {
        this.formulas.set(id, newFormula);
        this.size++;
    }
    getNombre() {
        return this.name;
    }

    getFormulas() {
        return this.formulas;

    }

    getSize() {
       return this.size;

    }
}