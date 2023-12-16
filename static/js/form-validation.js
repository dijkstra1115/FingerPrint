function checkFormInputs() {
    var allValid = true;
    var fingers = ['L1', 'L2', 'L3', 'L4', 'L5', 'R1', 'R2', 'R3', 'R4', 'R5'];

    for(var i = 0; i < fingers.length; i++) {
        // 检查代碼
        var codeElement = document.getElementById(fingers[i] + '-code');
        if(codeElement.value === '') {
            allValid = false;
            alert('請為 ' + fingers[i] + ' 選擇一個代碼。');
            break;
        }

        // 检查左量化
        var leftValueElement = document.getElementById(fingers[i] + '-left-value');
        if(leftValueElement.value === '') {
            allValid = false;
            alert('請為 ' + fingers[i] + ' 輸入左量化值。');
            break;
        }

        // 检查右量化
        var rightValueElement = document.getElementById(fingers[i] + '-right-value');
        if(rightValueElement.value === '') {
            allValid = false;
            alert('請為 ' + fingers[i] + ' 輸入右量化值。');
            break;
        }
    }

    // 检查报表名称
    var userNameElement = document.getElementById('user-name');
    if(userNameElement.value === '') {
        allValid = false;
        alert('請輸入報表名稱。');
    }

    // 檢查報表方案
    var pricingPlanElement = document.getElementById('pricing-plan');
    if(pricingPlanElement.value === '') {
        allValid = false;
        alert('請選擇報表方案。');
    }

    return allValid;
}

// 绑定到表单提交事件
document.getElementById('fingersDataForm').onsubmit = function(event) {
    if(!checkFormInputs()) {
        event.preventDefault(); // 阻止表单提交
    }
};
