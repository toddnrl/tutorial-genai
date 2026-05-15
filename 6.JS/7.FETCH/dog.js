const cateSelect = document.getElementById("cate");
const subCateSelect = document.getElementById("sub_cate");
const resultDiv = document.getElementById("result");

// 메인 품종 전체 불러오기
fetch("https://dog.ceo/api/breeds/list/all")
    .then(response => response.json())
    .then(data => {
        const breeds = data.message;

        for (let breed in breeds) {
            cateSelect.innerHTML += `
                <option value="${breed}">${breed}</option>
            `;
        }

        loadSubBreed();
    });

// 메인 품종 바뀌면 서브 품종 불러오기
cateSelect.addEventListener("change", loadSubBreed);

function loadSubBreed() {
    const category = cateSelect.value;

    fetch(`https://dog.ceo/api/breed/${category}/list`)
        .then(response => response.json())
        .then(data => {
            const subBreeds = data.message;

            subCateSelect.innerHTML = `
                <option value="">서브 품종 없음</option>
            `;

            subBreeds.forEach(sub => {
                subCateSelect.innerHTML += `
                    <option value="${sub}">${sub}</option>
                `;
            });
        });
}

// 버튼 클릭 시 이미지 요청
document.getElementById("cateBtn").addEventListener("click", () => {
    const category = cateSelect.value;
    const sub_category = subCateSelect.value;

    const url = sub_category
        ? `https://dog.ceo/api/breed/${category}/${sub_category}/images/random`
        : `https://dog.ceo/api/breed/${category}/images/random`;

    fetch(url)
        .then(response => response.json())
        .then(data => {
            resultDiv.innerHTML = `
                <h3>${category} ${sub_category}</h3>
                <img src="${data.message}" alt="강아지 사진">
            `;
        });
});