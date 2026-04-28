let time = 0;
let timer;
function start1(){
    clearInterval(timer);
    timer = setInterval(()=>{
        time++;
        let sec = Math.floor(time/10);
        let ms = time % 10;
        document.getElementById("watch").innerText = sec + "." + ms;
    }, 100);
}
function stop1(){
    clearInterval(timer);

    let sec = Math.floor(time / 10);
    let ms = time % 10;
    let resultTime = sec + "." + ms;

    let resultText = "";

    if(time == 100){
        resultText = "성공";
    }else{
        resultText = "실패";
    }

    // 👉 리스트에 추가
    let li = document.createElement("li");
    li.innerText = resultTime + "초 - " + resultText;

    document.getElementById("records").appendChild(li);

    // 실패하면 초기화
    if(resultText === "실패"){
        time = 0;
        document.getElementById("watch").innerText = "0.0";
    }
}