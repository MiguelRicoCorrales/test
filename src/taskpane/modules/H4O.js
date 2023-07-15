export const showErrorDialog = (text) => {
    console.log("text", text)

    //     const alertBox = document.createElement("div")
    //     alertBox.classList.add("alert")
    //     alertBox.innerText = text
    //     document.body.appendChild(alertBox)

    //     alertBox.addEventListener("click", () => alertBox.parentElement.removeChild(alertBox))
    // }
    // Office.context.ui.displayDialogAsync(
    //     'http://localhost:3000/errorDialog.html',
    //     { height: 200, width: 300 },
    //     function (result) {
    //         var dialog = result.value;
    //         dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (args) {
    //             if (args.message === 'closeDialog') {
    //                 dialog.close();
    //             }
    //         });
    //     }
    // );
};
