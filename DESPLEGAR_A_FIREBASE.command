#!/bin/zsh
cd "$(dirname "$0")"
echo "🚀 Iniciant desplegament a Firebase..."

# Intentar activar les APIs automàticament si gcloud està instal·lat
if command -v gcloud &> /dev/null; then
    echo "🔧 Activant APIs de Google Cloud..."
    gcloud services enable run.googleapis.com artifactregistry.googleapis.com --project infofoto-vector-art
fi

echo "📦 Pujant codi a Cloud Run..."
if command -v gcloud &> /dev/null; then
    gcloud run deploy infofoto-vector-service \
      --source . \
      --platform managed \
      --region europe-west1 \
      --allow-unauthenticated \
      --project infofoto-vector-art \
      --memory 2Gi \
      --timeout 300 \
      --set-env-vars PROJECT_ID=infofoto-vector-art,GOOGLE_CLIENT_ID=814718439112-2hcqqhsbbb2b67btpcqgtepakhmkhkkk.apps.googleusercontent.com
else
    echo "⚠️ No s'ha trobat gcloud. Pots instal·lar-lo o activar les APIs a la consola:"
    echo "👉 https://console.developers.google.com/apis/api/run.googleapis.com/overview?project=infofoto-vector-art"
fi

echo "🌐 Desplegant Hosting..."
firebase deploy --only hosting --project infofoto-vector-art

echo "✅ Procés finalitzat."
read -k 1 -s -r "?Prem qualsevol tecla per sortir..."
