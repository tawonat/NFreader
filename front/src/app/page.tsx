"use client";

import { useState } from "react";

export default function Home() {
  const [etapa, setEtapa] = useState(1);
  const [centroCusto, setCentroCusto] = useState("");
  const [arquivoNotas, setArquivoNotas] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);

  const avancarEtapa = () => {
    if (centroCusto.trim() === "") {
      alert("Por favor, insira o Centro de Custos Total.");
      return;
    }
    setEtapa(2);
  };

  const processarArquivos = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!arquivoNotas) {
      alert("Por favor, faça o upload do ZIP ou PDF das notas.");
      return;
    }

    setLoading(true);

    const formData = new FormData();
    formData.append("centroCusto", centroCusto);
    formData.append("notas", arquivoNotas);

    try {
      const response = await fetch("http://localhost:8000/processar", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const mensagem = await response.text();
        throw new Error(mensagem || "Erro na comunicação com o servidor.");
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.setAttribute("download", "Resultado_Notas.xlsx");
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);

      alert("Processamento e download concluídos com sucesso!");
    } catch (error) {
      console.error("Erro ao processar:", error);
      alert("Erro ao processar os arquivos. Verifique o backend e o formato enviado.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <main className="min-h-screen flex items-center justify-center bg-gray-50 p-4">
      <div className="bg-white p-8 rounded-lg shadow-md w-full max-w-md">
        <h1 className="text-2xl font-bold text-center text-gray-800 mb-6">
          Automação de Lançamento de Notas
        </h1>

        {etapa === 1 && (
          <div className="flex flex-col gap-4">
            <label className="text-gray-700 font-semibold">
              Qual o Centro de Custos Total?
            </label>
            <input
              type="text"
              className="border border-gray-300 p-2 rounded-md w-full text-black"
              placeholder="Ex: 102030 - Manutenção"
              value={centroCusto}
              onChange={(e) => setCentroCusto(e.target.value)}
            />
            <button
              onClick={avancarEtapa}
              className="bg-blue-600 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded-md transition"
            >
              Avançar
            </button>
          </div>
        )}

        {etapa === 2 && (
          <form onSubmit={processarArquivos} className="flex flex-col gap-6">
            <div className="bg-blue-50 p-4 rounded-md border border-blue-100">
              <p className="text-sm text-blue-800 font-medium mb-1">Centro de Custos:</p>
              <p className="text-gray-800 font-bold">{centroCusto}</p>
              <button
                type="button"
                onClick={() => setEtapa(1)}
                className="text-xs text-blue-600 underline mt-1"
              >
                Alterar Centro de Custos
              </button>
            </div>

            <div className="flex flex-col gap-2">
              <label className="text-gray-700 font-semibold">
                Upload das Notas Fiscais (PDF ou ZIP)
              </label>
              <input
                type="file"
                accept=".pdf, .zip"
                onChange={(e) => setArquivoNotas(e.target.files?.[0] || null)}
                className="border border-gray-300 p-2 rounded-md w-full text-gray-600 text-sm"
              />
              <p className="text-xs text-gray-500">
                O sistema utiliza automaticamente os cadastros fixos de Produtos e Serviços.
              </p>
            </div>

            <button
              type="submit"
              disabled={loading}
              className={`font-bold py-2 px-4 rounded-md transition text-white ${
                loading ? "bg-gray-400 cursor-not-allowed" : "bg-green-600 hover:bg-green-700"
              }`}
            >
              {loading ? "Processando e Analisando..." : "Gerar Planilha Final"}
            </button>
          </form>
        )}
      </div>
    </main>
  );
}
