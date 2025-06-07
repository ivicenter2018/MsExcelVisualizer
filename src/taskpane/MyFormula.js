// JavaScript source code



export default class MyFormula {

    name;
    body;   
    precedentsCells;
    

    constructor(nameAux, body) {
        this.name = nameAux;

        this.body = body;
       
        this.precedentsCells = new Set();
        
    }

    setPrecedentsCells(precedents){
       this.precedentsCells.add(precedents);
    }
    getNombre(){
        return this.name;
    }
    getBody(){
        return this.body;
    }

    
    getPrecedentsCells(){
        //console.log(this.precedentsCells);
        return this.precedentsCells;
    }

   
}
