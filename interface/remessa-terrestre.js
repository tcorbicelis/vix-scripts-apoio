const botaoRemessa = document.getElementById(
    "botao-remessa-terrestre"
);

if (!botaoRemessa) {
    console.error(
        "Botão 'botao-remessa-terrestre' não encontrado."
    );
    return;
}

botaoRemessa.addEventListener("click", async function () {

    // Evita múltiplos cliques
    botaoRemessa.disabled = true;

    const textoOriginal = botaoRemessa.innerText;
    botaoRemessa.innerText = "ABRINDO...";

    try {

        // Verifica se a API do PyWebView está disponível
        if (!window.pywebview || !window.pywebview.api) {

            throw new Error(
                "API do PyWebView não está disponível."
            );
        }

        // Chama o método Python:
        // VIX.executar_remessa()
        const resposta =
            await window.pywebview.api.executar_remessa();

        console.log("Resposta do Python:", resposta);

        if (resposta && resposta.sucesso) {

            console.log(
                "Remessa terrestre iniciada com sucesso."
            );

        } else {

            alert(
                resposta?.mensagem ||
                "Não foi possível iniciar a remessa terrestre."
            );
        }

    } catch (erro) {

        console.error(
            "Erro ao executar remessa terrestre:",
            erro
        );

        alert(
            "Erro ao executar a remessa terrestre:\n" +
            erro.message
        );

    } finally {

        // Restaura o botão
        botaoRemessa.disabled = false;
        botaoRemessa.innerText = textoOriginal;
    }

});