let count = 0;

const icb = document.getElementById("icb");
const dcb = document.getElementById("dcb");

icb.addEventListener("click", () => {
    count++;
    document.getElementById("result").innerText = count;
});
dcb.addEventListener("click", () => {
    count--;
    document.getElementById("result").innerText = count;
}); 