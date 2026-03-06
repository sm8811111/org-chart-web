const form = document.getElementById("upload-form");
const resultDiv = document.getElementById("result");

let scale = 1;
let posX = 0;
let posY = 0;

form.addEventListener("submit", async (e)=>{

    e.preventDefault();

    const fileInput = document.getElementById("file");

    if(!fileInput.files.length){
        alert("請選擇 Excel 檔案");
        return;
    }

    const formData = new FormData();
    formData.append("file",fileInput.files[0]);

    resultDiv.innerHTML="<p>生成中...</p>";

    const res = await fetch("/",{
        method:"POST",
        body:formData
    });

    const data = await res.json();

    if(data.img_path){

        resultDiv.innerHTML = `

        <div style="margin-bottom:10px;">

            <a href="${data.img_path}" download>
            <button>下載 PNG</button>
            </a>

            <a href="${data.svg_path}" download>
            <button>下載 SVG(超清晰)</button>
            </a>

        </div>

        <div id="viewer">

            <img id="orgImg"
            src="${data.img_path}?t=${Date.now()}">

        </div>

        `;

        const img = document.getElementById("orgImg");

        img.onload = ()=>{

            autoFit();
            enableDrag();
            enableWheelZoom();

        };

    }

});


function updateTransform(){

    const img = document.getElementById("orgImg");

    img.style.transform =
    `translate(${posX}px, ${posY}px) scale(${scale})`;

}



function autoFit(){

    const viewer = document.getElementById("viewer");
    const img = document.getElementById("orgImg");

    const vw = viewer.clientWidth;
    const vh = viewer.clientHeight;

    const iw = img.naturalWidth;
    const ih = img.naturalHeight;

    const scaleX = vw / iw;
    const scaleY = vh / ih;

    scale = Math.min(scaleX, scaleY);

    posX = (vw - iw * scale) / 2;
    posY = 0;

    updateTransform();

}



function enableWheelZoom(){

const viewer = document.getElementById("viewer");

viewer.addEventListener("wheel",function(e){

    e.preventDefault();

    const img = document.getElementById("orgImg");

    const rect = img.getBoundingClientRect();

    const mouseX = e.clientX - rect.left;
    const mouseY = e.clientY - rect.top;

    const zoom = e.deltaY < 0 ? 1.1 : 0.9;

    scale *= zoom;

    if(scale < 0.05) scale = 0.05;
    if(scale > 8) scale = 8;

    posX -= mouseX * (zoom - 1);
    posY -= mouseY * (zoom - 1);

    updateTransform();

});

}



function enableDrag(){

    const img = document.getElementById("orgImg");

    let dragging = false;
    let startX = 0;
    let startY = 0;

    img.addEventListener("mousedown", e => {

        dragging = true;

        // 記錄滑鼠與圖片的初始位置差
        startX = e.clientX - posX;
        startY = e.clientY - posY;

        img.style.cursor = "grabbing";

    });

    document.addEventListener("mouseup", e => {
        if(dragging){
            dragging = false;
            img.style.cursor = "grab";
        }
    });

    document.addEventListener("mousemove", e => {

        if(!dragging) return; // 只有按著滑鼠才移動

        posX = e.clientX - startX;
        posY = e.clientY - startY;

        updateTransform();
    });

    // 防止拖曳圖片時出現選取文字
    img.addEventListener("dragstart", e => e.preventDefault());

}