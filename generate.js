const pptxgen = require("pptxgenjs")
const fs = require('fs');
const { type } = require("os");

const response = fs.readFileSync('./dashboard.json', "utf-8")
const data = JSON.parse(response); 
const ppt = new pptxgen()


//General
const generalData = data.generalSection
const general = ppt.addSlide()
general.background = {color: "#F9F9F9"}

//Titulo
general.addShape(ppt.ShapeType.rect, { w:'28%', h:0.2, x:'36%', y:0.4    , fill: { color: "FAFF00" } });
general.addText("Dados Gerais", { y: 0.4,align: 'center', w: "100%",color: '#608CFE',fontFace: "Inter",fontSize: 32,bold: true,});

//Imagens
general.addImage({path: './logoBend.png', h:0.6, w:0.85, x:0.5, y:0.2})
general.addImage({path: './logoGamingIconBlack.png', h:0.6, w:0.6, x:9, y:0.2})

//Baras
general.addShape(ppt.ShapeType.rect, { w:'76%', h:0.01, x:'12%', y:'55%', fill: { color: "d9d9d9" } });
general.addShape(ppt.ShapeType.rect, { w: 0.01, h:'70%', x:'50%', y:'20%', fill: { color: "d9d9d9" } });

//MEDALHAS
general.addText("Medalhas", {align: 'center',w: '20%', x:'15%', y:'25%', color: '#608CFE', fontFace:"Inter" ,fontSize: 24 ,bold: true})
general.addText("Criadas", {align: 'center',w: '20%', x:'0%', y:'35%', color: '#545454', fontFace:"Inter" ,fontSize: 16 ,bold: true})
general.addText("Únicas \nConquistadas", {align: 'center', w: '20%', x:'15%', y:'35%', color: '#545454', fontFace:"Inter" ,fontSize: 16 ,bold: true})
general.addText("Total \nRepetidas", {align: 'center', w: '20%', x:'30%', y:'35%', color: '#545454', fontFace:"Inter" ,fontSize: 16 ,bold: true})
general.addText(String(generalData.medalInfo.created), {align: 'center', w: '20%',x:'0%', y:'45%', color: '#608CFE', fontFace:"Inter" ,fontSize: 16 ,bold: true})
general.addText(String(generalData.medalInfo.delivered.diff), {align: 'center', w: '20%',x:'15%', y:'45%', color: '#608CFE', fontFace:"Inter" ,fontSize: 16 ,bold: true})
general.addText(String(generalData.medalInfo.delivered.total), {align: 'center', w: '20%',x:'30%', y:'45%', color: '#608CFE', fontFace:"Inter" ,fontSize: 16 ,bold: true})

//Conquistas
general.addText("Conquistas", {align: 'center',w: '20%', x:'65%', y:'25%', color: '#608CFE', fontFace:"Inter" ,fontSize: 24 ,bold: true})
general.addText("Criadas", {align: 'center',w: '20%', x:'50%', y:'35%', color: '#545454', fontFace:"Inter" ,fontSize: 16 ,bold: true})
general.addText("Únicas \nConquistadas", {align: 'center', w: '20%', x:'65%', y:'35%', color: '#545454', fontFace:"Inter" ,fontSize: 16 ,bold: true})
general.addText("Total \nRepetidas", {align: 'center', w: '20%', x:'80%', y:'35%', color: '#545454', fontFace:"Inter" ,fontSize: 16 ,bold: true})
general.addText(String(generalData.awardInfo.created), {align: 'center', w: '20%',x:'50%', y:'45%', color: '#608CFE', fontFace:"Inter" ,fontSize: 16 ,bold: true})
general.addText(String(generalData.awardInfo.delivered), {align: 'center', w: '20%',x:'65%', y:'45%', color: '#608CFE', fontFace:"Inter" ,fontSize: 16 ,bold: true})
general.addText(String(generalData.awardInfo.delivered), {align: 'center', w: '20%',x:'80%', y:'45%', color: '#608CFE', fontFace:"Inter" ,fontSize: 16 ,bold: true})

//Quiz
general.addText("Quiz", {align: 'center',w: '20%', x:'15%', y:'65%', color: '#608CFE', fontFace:"Inter" ,fontSize: 24 ,bold: true})
general.addText("Criadas", {align: 'center',w: '20%', x:'0%', y:'75%', color: '#545454', fontFace:"Inter" ,fontSize: 16 ,bold: true})
general.addText("Únicas \nConquistadas", {align: 'center', w: '20%', x:'15%', y:'75%', color: '#545454', fontFace:"Inter" ,fontSize: 16 ,bold: true})
general.addText("Total \nRepetidas", {align: 'center', w: '20%', x:'30%', y:'75%', color: '#545454', fontFace:"Inter" ,fontSize: 16 ,bold: true})
general.addText(String(generalData.quizInfo.created), {align: 'center', w: '20%',x:'0%', y:'85%', color: '#608CFE', fontFace:"Inter" ,fontSize: 16 ,bold: true})
general.addText(String(generalData.quizInfo.played), {align: 'center', w: '20%',x:'15%', y:'85%', color: '#608CFE', fontFace:"Inter" ,fontSize: 16 ,bold: true})
general.addText(String(generalData.quizInfo.completed), {align: 'center', w: '20%',x:'30%', y:'85%', color: '#608CFE', fontFace:"Inter" ,fontSize: 16 ,bold: true})

//Loja de recompensas
general.addText("Loja de recompensas", {align: 'center', w: '50%', x:'50%', y:'65%', color: '#608CFE', fontFace:"Inter" ,fontSize: 24 ,bold: true})
general.addText("Criadas", {align: 'center',w: '20%', x:'50%', y:'75%', color: '#545454', fontFace:"Inter" ,fontSize: 16 ,bold: true})
general.addText("Únicas \nConquistadas", {align: 'center', w: '20%', x:'65%', y:'75%', color: '#545454', fontFace:"Inter" ,fontSize: 16 ,bold: true})
general.addText("Total \nRepetidas", {align: 'center', w: '20%', x:'80%', y:'75%', color: '#545454', fontFace:"Inter" ,fontSize: 16 ,bold: true})
general.addText(String(generalData.rewardInfo.created), {align: 'center', w: '20%',x:'50%', y:'85%', color: '#608CFE', fontFace:"Inter" ,fontSize: 16 ,bold: true})
general.addText(String(generalData.rewardInfo.bought), {align: 'center', w: '20%',x:'65%', y:'85%', color: '#608CFE', fontFace:"Inter" ,fontSize: 16 ,bold: true})
general.addText(String(generalData.rewardInfo.paid), {align: 'center', w: '20%',x:'80%', y:'85%', color: '#608CFE', fontFace:"Inter" ,fontSize: 16 ,bold: true})

//Medalhas
const medalData = data.medalSession
const medal = ppt.addSlide()
medal.background = {color: "#F9F9F9"}

//Titulo
medal.addShape(ppt.ShapeType.rect, { w:'20%', h:0.2, x:'40%', y:0.4    , fill: { color: "FAFF00" } });
medal.addText("Medalhas", { y: 0.4,align: 'center', w: "100%",color: '#608CFE',fontFace: "Inter",fontSize: 32,bold: true,});

//Imagens
medal.addImage({path: './logoBend.png', h:0.6, w:0.85, x:0.5, y:0.2})
medal.addImage({path: './logoGamingIconBlack.png', h:0.6, w:0.6, x:9, y:0.2})

//Titulos Tabela 
medal.addText("Descrição", {align: 'left',w:'30%', x:'4%', y:'30%', color: '#608CFE', fontFace:"Inter" ,fontSize: 18, bold: true})
medal.addShape(ppt.ShapeType.rect, { w: 0.01, h:'10%', x:'37%', y:'25%', fill: { color: "d9d9d9" } });
medal.addText("Tipo", {align: 'center', w:'20%', x:'40%', y:'30%', color: '#608CFE', fontFace:"Inter" ,fontSize: 18, bold: true})
medal.addShape(ppt.ShapeType.rect, { w: 0.01, h:'10%', x:'63%', y:'25%', fill: { color: "d9d9d9" } });
medal.addText("Pontos",{align: 'center', w:'12%', x:'63%', y:'30%', color: '#608CFE', fontFace:"Inter" ,fontSize: 18, bold: true})
medal.addShape(ppt.ShapeType.rect, { w: 0.01, h:'10%', x:'75%', y:'25%', fill: { color: "d9d9d9" } });
medal.addText("Total \nGanho", {align: 'center',w:'12%', x:'75%', y:'30%', color: '#608CFE', fontFace:"Inter" ,fontSize: 18, bold: true})
medal.addShape(ppt.ShapeType.rect, { w: 0.01, h:'10%', x:'87%', y:'25%', fill: { color: "d9d9d9" } });
medal.addText("Total de \nusuários", {align: 'center',w:'12%', x:'87%', y:'30%', color: '#608CFE', fontFace:"Inter" ,fontSize: 20, bold: true})

let yDataIncrement = 0
for(item of medalData){
    yDataIncrement += 10
    medal.addText(item.name, {align: 'left',w:'30%', x:'4%', y:`${30+yDataIncrement}%`, color: '#545454', fontFace:"Inter" ,fontSize: 16 ,bold: true})
    medal.addShape(ppt.ShapeType.rect, { w: 0.01, h:'12%', x:'37%', y:`${25+yDataIncrement}%`, fill: { color: "d9d9d9" } });
    medal.addText(item.type, {align: 'center', w:'20%', x:'40%', y:`${30+yDataIncrement}%`, color: '#0e3db5', fontFace:"Inter" ,fontSize: 16, bold: true,})
    medal.addShape(ppt.ShapeType.rect, { w: 0.01, h:'12%', x:'63%', y:`${25+yDataIncrement}%`, fill: { color: "d9d9d9" } });
    medal.addText(String(item.points),{align: 'center', w:'12%', x:'63%', y:`${30+yDataIncrement}%`, color: '#545454', fontFace:"Inter" ,fontSize: 16 ,bold: true})
    medal.addShape(ppt.ShapeType.rect, { w: 0.01, h:'12%', x:'75%', y:`${25+yDataIncrement}%`, fill: { color: "d9d9d9" } });
    medal.addText(String(item.totalQuantity), {align: 'center',w:'12%', x:'75%', y:`${30+yDataIncrement}%`, color: '#545454', fontFace:"Inter" ,fontSize: 16 ,bold: true})
    medal.addShape(ppt.ShapeType.rect, { w: 0.01, h:'12%', x:'87%', y:`${25+yDataIncrement}%`, fill: { color: "d9d9d9" } });
    medal.addText(String(item.usersQuantity), {align: 'center',w:'12%', x:'87%', y:`${30+yDataIncrement}%`, color: '#545454', fontFace:"Inter" ,fontSize: 16 ,bold: true})
}


ppt.writeFile({ fileName: 'Apresentação.pptx' })
.then(() => { console.log("Arquivo PPTX Gerado!")})

