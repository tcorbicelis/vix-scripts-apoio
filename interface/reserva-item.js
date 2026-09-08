const tabelaBody = document.getElementById("tabela-body");

        const quantidadeLinhas = 50;

        for (let i = 0; i < quantidadeLinhas; i++) {

            const linha = document.createElement("tr");

            linha.innerHTML = `
                <td>
                    <input
                        type="text"
                        name="reserva[]"
                        placeholder="Reserva"
                    >
                </td>

                <td>
                    <input
                        type="text"
                        name="item[]"
                        placeholder="Item"
                    >
                </td>
            `;

            tabelaBody.appendChild(linha);
        }