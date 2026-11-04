# FoundryInsights container: PowerShell 7 + Python (stats/Excel) + Azure CLI.
# Build:  docker build -t foundry-insights .
# Run (Azure CLI login inside the container, results written to ./output on the host):
#   docker run --rm -it -v "${PWD}/output:/app/output" foundry-insights \
#     -c "az login --use-device-code --tenant <tenant> && pwsh -File scripts/Run-FoundryInsights.ps1 -TenantId <tenant> -Days 30"
# Managed identity (e.g. Azure Container Apps job / ACI with an identity):
#   pwsh -File scripts/Run-FoundryInsights.ps1 -TenantId <tenant> -AuthMode ManagedIdentity
FROM mcr.microsoft.com/powershell:lts-ubuntu-22.04

ENV DEBIAN_FRONTEND=noninteractive PIP_NO_CACHE_DIR=1 PYTHONDONTWRITEBYTECODE=1
RUN apt-get update \
 && apt-get install -y --no-install-recommends python3 python3-pip ca-certificates curl \
 && curl -sL https://aka.ms/InstallAzureCLIDeb | bash \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
# Optional private/corporate package index: docker build --build-arg PIP_INDEX_URL=https://... .
ARG PIP_INDEX_URL
RUN if [ -n "$PIP_INDEX_URL" ]; then export PIP_INDEX_URL; else unset PIP_INDEX_URL; fi \
 && python3 -m pip install -r requirements.txt
COPY scripts/ scripts/
COPY workbooks/ workbooks/
COPY prices.example.csv README.md ./

RUN useradd --create-home fi && mkdir -p /app/output && chown -R fi /app
USER fi
ENTRYPOINT ["/bin/bash"]
CMD ["-c", "pwsh -File scripts/Run-FoundryInsights.ps1 -?"]
