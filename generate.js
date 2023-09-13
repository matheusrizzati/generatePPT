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
general.addText("Dados Gerais", { y: 0.4,align: 'center', w: "100%",color: '#608CFE',fontFace: "Inter",fontSize: 32,bold: true,});

//Imagens
general.addImage({path: './images/logoBend.png', h:0.6, w:0.85, x:0.5, y:0.2})
general.addImage({path: './images/logoGamingIconBlack.png', h:0.6, w:0.6, x:9, y:0.2})

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
medal.addText("Medalhas", { y: 0.4,align: 'center', w: "100%",color: '#608CFE',fontFace: "Inter",fontSize: 32,bold: true,});

//Imagens
medal.addImage({path: './images/logoBend.png', h:0.6, w:0.85, x:0.5, y:0.2})
medal.addImage({path: './images/logoGamingIconBlack.png', h:0.6, w:0.6, x:9, y:0.2})

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


//conquistas
const awardData = data.awardSession
const award = ppt.addSlide()
award.background = {color: "#F9F9F9"}

//Titulo
award.addText("Conquistas", { y: 0.4,align: 'center', w: "100%",color: '#608CFE',fontFace: "Inter",fontSize: 32, bold: true});
award.addImage({path: './images/logoBend.png', h:0.6, w:0.85, x:0.5, y:0.2})
award.addImage({path: './images/logoGamingIconBlack.png', h:0.6, w:0.6, x:9, y:0.2})
//Titulos tabela
award.addText("Pontos", {align: 'center', w:'20%', x:'40%', y:'30%', color: '#608CFE', fontFace:"Inter" ,fontSize: 18})
award.addText("Total de \nusuários", {align: 'center', w:'20%', x:'70%', y:'30%', color: '#545454', fontFace:"Inter" ,fontSize: 18})

//conteudo
let wardYacc = 0
for(item of awardData){
    wardYacc += 10
    award.addText(item.name, {align: 'left', w:'30%', x:'10%', y:`${30+wardYacc}%`, color: '#545454', fontFace:"Inter" ,fontSize: 18, bold: true})
    award.addText(String(item.points), {align: 'center', w:'20%', x:'40%', y:`${30+wardYacc}%`, color: '#608CFE', fontFace:"Inter" ,fontSize: 18, bold: true})
    award.addText(String(item.usersQuantity), {align: 'center', w:'20%', x:'70%', y:`${30+wardYacc}%`, color: '#545454', fontFace:"Inter" ,fontSize: 18, bold: true})
}

//quizRecorde
const quizData = data.quizSession
const quizRecorde = ppt.addSlide()
quizRecorde.background = {color: "#F9F9F9"}

//Titulo
quizRecorde.addText("Quiz", { y: 0.4,align: 'center', w: "100%",color: '#608CFE',fontFace: "Inter",fontSize: 32, bold: true});
quizRecorde.addImage({path: './images/logoBend.png', h:0.6, w:0.85, x:0.5, y:0.2})
quizRecorde.addImage({path: './images/logoGamingIconBlack.png', h:0.6, w:0.6, x:9, y:0.2})
//TITULO2
quizRecorde.addText("TIPO", {align: 'left', w:'20%', x:'5%', y:`25%`, color: '#545454', fontFace:"Inter" ,fontSize: 18})
quizRecorde.addText("RECORD", {align: 'left', w:'30%', x:'5%', y:`30%`, color: '#608CFE', fontFace:"Inter" ,fontSize: 32, bold: true})
quizRecorde.addShape(ppt.ShapeType.rect, {align:'left', w:'76%', h:0.01, x:'6%', y:'35%', fill: { color: "545454" } });
quizRecorde.addText("Pontos", {align: 'center', w:'12%', x:'26%', y:'40%', color: '#608CFE', fontFace:"Inter", fontSize: 16})
quizRecorde.addText("Perguntas", {align: 'center', w:'12%', x:'38%', y:'40%', color: '#545454', fontFace:"Inter", fontSize: 16})
quizRecorde.addText("Tentativas", {align: 'center', w:'12%', x:'50%', y:'40%', color: '#608CFE', fontFace:"Inter", fontSize: 16})

const recordData = Object.values(quizData).filter(item => item.type === 'record');
let recordAcc = 0
for(item of recordData){
    recordAcc += 10
    quizRecorde.addText(item.name, {align: 'left', w:'23%', x:'5%', y:`${40+recordAcc}%`, color: '#545454', fontFace:"Inter", fontSize: 16, bold: true})
    quizRecorde.addText(Number(item.points), {align: 'center', w:'12%', x:'26%', y:`${40+recordAcc}%`, color: '#608CFE', fontFace:"Inter", fontSize: 16, bold: true})
    quizRecorde.addText(Number(item.questionsQuantity), {align: 'center', w:'12%', x:'38%', y:`${40+recordAcc}%`, color: '#545454', fontFace:"Inter", fontSize: 16, bold: true})
    quizRecorde.addText("?", {align: 'center', w:'12%', x:'50%', y:`${40+recordAcc}%`, color: '#608CFE', fontFace:"Inter", fontSize: 16, bold: true})
}
//Quiz certificado
const quizCertificado = ppt.addSlide()
quizCertificado.background = {color: "#F9F9F9"}

//Titulo
quizCertificado.addText("Quiz", { y: 0.4,align: 'center', w: "100%",color: '#608CFE',fontFace: "Inter",fontSize: 32, bold: true});
quizCertificado.addImage({path: './images/logoBend.png', h:0.6, w:0.85, x:0.5, y:0.2})
quizCertificado.addImage({path: './images/logoGamingIconBlack.png', h:0.6, w:0.6, x:9, y:0.2})

//TITULO2
quizCertificado.addText("TIPO", {align: 'left', w:'20%', x:'5%', y:`25%`, color: '#545454', fontFace:"Inter" ,fontSize: 18})
quizCertificado.addText("CERTIFICADO", {align: 'left', w:'30%', x:'5%', y:`30%`, color: '#608CFE', fontFace:"Inter" ,fontSize: 32, bold: true})
quizCertificado.addShape(ppt.ShapeType.rect, {align:'left', w:'76%', h:0.01, x:'6%', y:'35%', fill: { color: "545454" } });
quizCertificado.addText("Pontos", {align: 'center', w:'12%', x:'26%', y:'40%', color: '#608CFE', fontFace:"Inter", fontSize: 16})
quizCertificado.addText("Perguntas", {align: 'center', w:'12%', x:'38%', y:'40%', color: '#545454', fontFace:"Inter", fontSize: 16})
quizCertificado.addText("Tentativas", {align: 'center', w:'12%', x:'50%', y:'40%', color: '#608CFE', fontFace:"Inter", fontSize: 16})
quizCertificado.addText("N. de Corte", {align: 'center', w:'15%', x:'62%', y:'40%', color: '#545454', fontFace:"Inter", fontSize: 16})
quizCertificado.addText("Certificados", {align: 'center', w:'15%', x:'77%', y:'40%', color: '#608CFE', fontFace:"Inter", fontSize: 16})

const certificadoData = Object.values(quizData).filter(item => item.type === 'certificate');
let certificateAcc = 0
for(item of certificadoData){
    certificateAcc += 10
    quizCertificado.addText(item.name, {align: 'left', w:'23%', x:'5%', y:`${40+certificateAcc}%`, color: '#545454', fontFace:"Inter", fontSize: 16, bold: true})
    quizCertificado.addText(Number(item.points), {align: 'center', w:'12%', x:'26%', y:`${40+certificateAcc}%`, color: '#608CFE', fontFace:"Inter", fontSize: 16, bold: true})
    quizCertificado.addText(Number(item.questionsQuantity), {align: 'center', w:'12%', x:'38%', y:`${40+certificateAcc}%`, color: '#545454', fontFace:"Inter", fontSize: 16, bold: true})
    quizCertificado.addText("?", {align: 'center', w:'12%', x:'50%', y:`${40+certificateAcc}%`, color: '#608CFE', fontFace:"Inter", fontSize: 16, bold: true})
    quizCertificado.addText("?", {align: 'center', w:'15%', x:'62%', y:`${40+certificateAcc}%`, color: '#608CFE', fontFace:"Inter", fontSize: 16, bold: true})
    quizCertificado.addText("?", {align: 'center', w:'15%', x:'77%', y:`${40+certificateAcc}%`, color: '#608CFE', fontFace:"Inter", fontSize: 16, bold: true})
} 


//Recompensas
const rewardData = data.rewardSession
const reward = ppt.addSlide()
reward.background = {color: "#F9F9F9"}

//Titulo
reward.addText("Recompensas", { y: 0.4,align: 'center', w: "100%",color: '#608CFE',fontFace: "Inter",fontSize: 32,bold: true,});

//Imagens
reward.addImage({path: './images/logoBend.png', h:0.6, w:0.85, x:0.5, y:0.2})
reward.addImage({path: './images/logoGamingIconBlack.png', h:0.6, w:0.6, x:9, y:0.2})

//Grades
reward.addShape(ppt.ShapeType.rect, { w: 0.01, h:'10%', x:'26%', y:'30%', fill: { color: "d9d9d9" } });
reward.addText("Preço", {align: 'center', w:'18%', x:'26%', y:'35%', color:'#608CFE', fontFace:'Inter', fontSize:18})
reward.addShape(ppt.ShapeType.rect, { w: 0.01, h:'10%', x:'44%', y:'30%', fill: { color: "d9d9d9" } });
reward.addText("Disponíveis", {align: 'center', w:'18%', x:'44%', y:'35%', color:'#545454', fontFace:'Inter', fontSize:18})
reward.addShape(ppt.ShapeType.rect, { w: 0.01, h:'10%', x:'62%', y:'30%', fill: { color: "d9d9d9" } });
reward.addText("Pagamentos \nem aberto", {align: 'center', w:'18%', x:'62%', y:'35%', color:'#608CFE', fontFace:'Inter', fontSize:18})
reward.addShape(ppt.ShapeType.rect, { w: 0.01, h:'10%', x:'80%', y:'30%', fill: { color: "d9d9d9" } });
reward.addText("Pagamentos \nfechados", {align: 'center', w:'18%', x:'80%', y:'35%', color:'#545454', fontFace:'Inter', fontSize:18})

let rewardAcc = 0
for(item of rewardData){
    rewardAcc += 10
reward.addText(item.name, {align: 'left', w:'24%', x:'2%', y:`${35+rewardAcc}%`, color:'#545454', fontFace:'Inter', fontSize:18, bold: true})
reward.addShape(ppt.ShapeType.rect, { w: 0.01, h:'10%', x:'26%', y:`${30+rewardAcc}%`, fill: { color: "d9d9d9" } });
reward.addText(String(item.pointsPrice), {align: 'center', w:'18%', x:'26%', y:`${35+rewardAcc}%`, color:'#608CFE', fontFace:'Inter', fontSize:18, bold: true})
reward.addShape(ppt.ShapeType.rect, { w: 0.01, h:'10%', x:'44%', y:`${30+rewardAcc}%`, fill: { color: "d9d9d9" } });
reward.addText(String(item.availableQuantity), {align: 'center', w:'18%', x:'44%', y:`${35+rewardAcc}%`, color:'#545454', fontFace:'Inter', fontSize:18, bold: true})
reward.addShape(ppt.ShapeType.rect, { w: 0.01, h:'10%', x:'62%', y:`${30+rewardAcc}%`, fill: { color: "d9d9d9" } });
reward.addText(String(item.payments.opened.length), {align: 'center', w:'18%', x:'62%', y:`${35+rewardAcc}%`, color:'#608CFE', fontFace:'Inter', fontSize:18, bold: true})
reward.addShape(ppt.ShapeType.rect, { w: 0.01, h:'10%', x:'80%', y:`${30+rewardAcc}%`, fill: { color: "d9d9d9" } });
reward.addText(String(item.payments.closed.length), {align: 'center', w:'18%', x:'80%', y:`${35+rewardAcc}%`, color:'#545454', fontFace:'Inter', fontSize:18, bold: true})   
}


//Recompensas
const paymentData = data.paymentSession
const payment = ppt.addSlide()
payment.background = {color: "#F9F9F9"}

//Titulo
payment.addText("Pagamentos", { y: 0.4,align: 'center', w: "100%",color: '#608CFE',fontFace: "Inter",fontSize: 32,bold: true,});

//Imagens
payment.addImage({path: './images/logoBend.png', h:0.6, w:0.85, x:0.5, y:0.2})
payment.addImage({path: './images/logoGamingIconBlack.png', h:0.6, w:0.6, x:9, y:0.2})

//TEXT
payment.addText('Em aberto', {align:'center', w:'20%', x:'20%', y:'45%', color:'#608CFE', fontFace:'Inter', fontSize:24, bold:true})
payment.addText(String(paymentData.opened.payments.length), {align:'center', w:'20%', x:'20%', y:'55%', color:'#545454', fontFace:'Inter', fontSize:24, bold:true})

payment.addText('Fechado', {align:'center', w:'20%', x:'60%', y:'45%', color:'#608CFE', fontFace:'Inter', fontSize:24, bold:true})
payment.addText(String(paymentData.closed.payments.length), {align:'center', w:'20%', x:'60%', y:'55%', color:'#545454', fontFace:'Inter', fontSize:24, bold:true})


//

//Recompensas
const taskData = data.taskSession
const task = ppt.addSlide()
task.background = {color: "#F9F9F9"}

//Titulo
task.addText("Tarefas", { y: 0.4,align: 'center', w: "100%",color: '#608CFE',fontFace: "Inter",fontSize: 32,bold: true,});

//Imagens
task.addImage({path: './images/logoBend.png', h:0.6, w:0.85, x:0.5, y:0.2})
task.addImage({path: './images/logoGamingIconBlack.png', h:0.6, w:0.6, x:9, y:0.2})

//TEXT
task.addText('Em aberto', {align:'center', w:'20%', x:'15%', y:'40%', color:'#545454', fontFace:'Inter', fontSize:24, bold:true})
task.addText('Simples', {align:'center', w:'12%', x:'7%', y:'50%', color:'#608CFE', fontFace:'Inter', fontSize:16,})
task.addText('Missão', {align:'center', w:'12%', x:'19%', y:'50%', color:'#608CFE', fontFace:'Inter', fontSize:16,})
task.addText('Periódica', {align:'center', w:'12%', x:'31%', y:'50%', color:'#608CFE', fontFace:'Inter', fontSize:16,})
task.addText(String(taskData[0].opened.simpleTasks.length), {align:'center', w:'12%', x:'7%', y:'60%', color:'#545454', fontFace:'Inter', fontSize:16, bold:true})
task.addText(String(taskData[0].opened.missionTasks.length), {align:'center', w:'12%', x:'19%', y:'60%', color:'#545454', fontFace:'Inter', fontSize:16, bold:true})
task.addText(String(taskData[0].opened.periodicTasks.length), {align:'center', w:'12%', x:'31%', y:'60%', color:'#545454', fontFace:'Inter', fontSize:16, bold:true})

task.addText('Fechada', {align:'center', w:'20%', x:'65%', y:'40%', color:'#545454', fontFace:'Inter', fontSize:24, bold:true})
task.addText('Simples', {align:'center', w:'12%', x:'57%', y:'50%', color:'#608CFE', fontFace:'Inter', fontSize:16,})
task.addText('Missão', {align:'center', w:'12%', x:'69%', y:'50%', color:'#608CFE', fontFace:'Inter', fontSize:16,})
task.addText('Periódica', {align:'center', w:'12%', x:'81%', y:'50%', color:'#608CFE', fontFace:'Inter', fontSize:16,})
task.addText(String(taskData[0].closed.simpleTasks.length), {align:'center', w:'12%', x:'57%', y:'60%', color:'#545454', fontFace:'Inter', fontSize:16, bold:true})
task.addText(String(taskData[0].closed.missionTasks.length), {align:'center', w:'12%', x:'69%', y:'60%', color:'#545454', fontFace:'Inter', fontSize:16, bold:true})
task.addText(String(taskData[0].closed.periodicTasks.length), {align:'center', w:'12%', x:'81%', y:'60%', color:'#545454', fontFace:'Inter', fontSize:16, bold:true})


/////
const userData = data.userSession
for(item of userData){
const user = ppt.addSlide()
user.background = {color: "#F9F9F9"}
//Titulo
user.addText("Usuários", { y: 0.4,align: 'center', w: "100%",color: '#608CFE',fontFace: "Inter",fontSize: 32,bold: true,});
//Imagens
user.addImage({path: './images/logoBend.png', h:0.6, w:0.85, x:0.5, y:0.2})
user.addImage({path: './images/logoGamingIconBlack.png', h:0.6, w:0.6, x:9, y:0.2})

user.addText(item.name, {align: 'left', w:'30%', x:'5%', y:'25%', color:'#608CFE', fontSize:24, bold:true})
user.addText(
        (item.level !== undefined) ? item.level.name : 'SEM NÍVEL'
    ,{align: 'left', w:'30%', x:'5%', y:'30%', color:'#545454', fontSize:16, bold:false})
user.addText("Saldo", {align: 'left', w:'20%', x:'40%', y:'25%', color:'#545454', fontSize:18, bold:false})
user.addText(String(item.points), {align: 'left', w:'20%', x:'40%', y:'30%', color:'#608CFE', fontSize:18, bold:true})
user.addText("Acumulado", {align: 'left', w:'20%', x:'60%', y:'25%', color:'#545454', fontSize:18, bold:false})
user.addText(String(item.acumulatedPoints), {align: 'left', w:'20%', x:'60%', y:'30%', color:'#608CFE', fontSize:18, bold:true})

user.addText("Medalhas", {align:'center', w:'20%', x:'5%', y:'50%', color:'#608CFE', fontSize:20, bold:true})
user.addText("Total", {align:'center', w:'20%', x:'5%', y:'60%', color:'#545454', fontSize:18, bold:false})
user.addText(String(item.totalMedals), {align:'center', w:'20%', x:'5%', y:'65%', color:'#608CFE', fontSize:18, bold:true})
user.addText("Únicas", {align:'center', w:'20%', x:'5%', y:'75%', color:'#545454', fontSize:18, bold:false})
user.addText(`${String(item.medals.length)}/${String(item.groupMedals)}`, {align:'center', w:'20%', x:'5%', y:'80%', color:'#608CFE', fontSize:18, bold:true})
user.addShape(ppt.ShapeType.rect, { w:'16%', h:'2%', x:'7%', y:'85%', fill: { color: "949494" } });
user.addShape(ppt.ShapeType.rect, { w:`${16 / item.groupMedals * item.medals.length}%`, h:'2%', x:'7%', y:'85%', fill: { color: "608CFE" } });


user.addText("Pagamentos", {align:'center', w:'20%', x:'40%', y:'50%', color:'#608CFE', fontSize:20, bold:true})
user.addText("A receber", {align:'center', w:'20%', x:'30%', y:'58%', color:'#608CFE', fontSize:18, bold:false})
user.addText(`${String(item.payments.toReceive.opened.length)} em aberto`, {align:'center', w:'20%', x:'30%', y:'64%', color:'#545454', fontSize:18, bold:false})
user.addText(`${String(item.payments.toReceive.closed.length)} fechados`, {align:'center', w:'20%', x:'30%', y:'70%', color:'#545454', fontSize:18, bold:false})
user.addText("A pagar", {align:'center', w:'20%', x:'50%', y:'58%', color:'#608CFE', fontSize:18, bold:false})
user.addText(`${String(item.payments.toPay.opened.length)} em aberto`, {align:'center', w:'20%', x:'50%', y:'64%', color:'#545454', fontSize:18, bold:false})
user.addText(`${String(item.payments.toPay.closed.length)} fechados`, {align:'center', w:'20%', x:'50%', y:'70%', color:'#545454', fontSize:18, bold:false})

user.addText("Quiz", {align:'center', w:'20%', x:'75%', y:'50%', color:'#608CFE', fontSize:20, bold:true})
user.addText("Certificações", {align:'center', w:'20%', x:'75%', y:'58%', color:'#545454', fontSize:18, bold:false})
user.addText(String(item.quiz.certifiedQuizzes.length), {align:'center', w:'20%', x:'75%', y:'63%', color:'#608CFE', fontSize:18, bold:true})
user.addText("Disponíveis", {align:'center', w:'20%', x:'75%', y:'71%', color:'#545454', fontSize:18, bold:false})
user.addText(String(item.quiz.availableQuizzes.length), {align:'center', w:'20%', x:'75%', y:'76%', color:'#608CFE', fontSize:18, bold:true})
user.addText("Respondidos", {align:'center', w:'20%', x:'75%', y:'84%', color:'#545454', fontSize:18, bold:false})
user.addText(String(item.quiz.answeredQuizzes.length), {align:'center', w:'20%', x:'75%', y:'89%', color:'#608CFE', fontSize:18, bold:true})

////
const userTasks = ppt.addSlide()
userTasks.background = {color: "#F9F9F9"}
//Titulo
userTasks.addText("Usuários", { y: 0.4,align: 'center', w: "100%",color: '#608CFE',fontFace: "Inter",fontSize: 32,bold: true,});
//Imagens
userTasks.addImage({path: './images/logoBend.png', h:0.6, w:0.85, x:0.5, y:0.2})
userTasks.addImage({path: './images/logoGamingIconBlack.png', h:0.6, w:0.6, x:9, y:0.2})

userTasks.addText(item.name, {align: 'left', w:'30%', x:'5%', y:'25%', color:'#608CFE', fontSize:24, bold:true})
userTasks.addText(
    (item.level !== undefined) ? item.level.name : 'SEM NÍVEL'
    , {align: 'left', w:'30%', x:'5%', y:'30%', color:'#545454', fontSize:16, bold:false})
userTasks.addText("Saldo", {align: 'left', w:'20%', x:'40%', y:'25%', color:'#545454', fontSize:18, bold:false})
userTasks.addText(String(item.points), {align: 'left', w:'20%', x:'40%', y:'30%', color:'#608CFE', fontSize:18, bold:true})
userTasks.addText("Acumulado", {align: 'left', w:'20%', x:'60%', y:'25%', color:'#545454', fontSize:18, bold:false})
userTasks.addText(String(item.acumulatedPoints), {align: 'left', w:'20%', x:'60%', y:'30%', color:'#608CFE', fontSize:18, bold:true})

userTasks.addText("Tarefas", {align:'center', w:'20%', x:'40%', y:'52%', color:"#608CFE", fontSize:20, bold:true})
userTasks.addText("Dono", {align:'center', w:'20%', x:'20%', y:'60%', color:"#608CFE", fontSize:18, bold:false})
userTasks.addText(`${String(item.tasks.owners.opened.length)} em aberto`, {align:'center', w:'20%', x:'20%', y:'65%', color:"#545454", fontSize:18, bold:false})
userTasks.addText(`${String(item.tasks.owners.closed.length )} fechados`, {align:'center', w:'20%', x:'20%', y:'70%', color:"#545454", fontSize:18, bold:false})
userTasks.addText("Observador", {align:'center', w:'20%', x:'40%', y:'60%', color:"#608CFE", fontSize:18, bold:false})
userTasks.addText(`${String(item.tasks.observers.opened.length)} em aberto`, {align:'center', w:'20%', x:'40%', y:'65%', color:"#545454", fontSize:18, bold:false})
userTasks.addText(`${String(item.tasks.observers.closed.length )} fechados`, {align:'center', w:'20%', x:'40%', y:'70%', color:"#545454", fontSize:18, bold:false})
userTasks.addText("Aprovador", {align:'center', w:'20%', x:'60%', y:'60%', color:"#608CFE", fontSize:18, bold:false})
userTasks.addText(`${String(item.tasks.approvers.opened.length)} em aberto`, {align:'center', w:'20%', x:'60%', y:'65%', color:"#545454", fontSize:18, bold:false})
userTasks.addText(`${String(item.tasks.approvers.closed.length)} fechados`, {align:'center', w:'20%', x:'60%', y:'70%', color:"#545454", fontSize:18, bold:false})

userTasks.addText("Conquistas", {align:'center', w:'20%', x:'40%', y:'85%', color:'#608CFE', fontSize:20, bold:true})
userTasks.addText(String(item.awards.length), {align:'right', w:'10%', x:'30%', y:'90%', color:'#545454', fontSize:18, bold:true})
userTasks.addShape(ppt.ShapeType.rect, { w:'20%', h:'2%', x:'40%', y:'90%', fill: { color: "949494" } });
userTasks.addShape(ppt.ShapeType.rect, { w:`${20 / item.groupAwards * item.awards.length}%`, h:'2%', x:'40%', y:'90%', fill: { color: "608CFE" } });
userTasks.addText(String(item.groupAwards), {align:'left', w:'10%', x:'60%', y:'90%', color:'#545454', fontSize:18, bold:true})
}

ppt.writeFile({ fileName: `Relatório_Dados_${new Date(Date.now()).getFullYear()}${new Date(Date.now()).getMonth() + 1}${new Date(Date.now()).getDate()}.pptx` })
.then(() => { console.log("Arquivo PPTX Gerado!")})

