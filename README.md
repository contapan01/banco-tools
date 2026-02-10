# Portal de Herramientas Bancarias 🏦

Este repositorio contiene una aplicación web desarrollada con **Streamlit** para procesar movimientos de cuentas bancarias y exportarlos al formato de **Zoho Books**.

## 🚀 Características
- **Multi-Banco**: Soporte para Banco BGP y Motor Bank (MB).
- **Consolidación Inteligente**: Lógica para unir movimientos relacionados en Motor Bank.
- **Resumen Automático**: Visualización de totales de débitos y créditos.
- **Docker Ready**: Listo para desplegar en Hetzner vía Easypanel o Docker Compose.

## 🛠️ Instalación Local
1. `pip install -r requirements.txt`
2. `streamlit run app.py`

## 🐳 Despliegue con Easypanel
Simplemente conecta este repositorio a un nuevo proyecto en tu Easypanel. El `Dockerfile` se encargará de todo.
