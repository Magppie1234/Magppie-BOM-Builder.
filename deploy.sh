#!/bin/bash
set -e

echo "=== Vercel Auto Deployment Script ==="

# 1. Login using npx vercel
echo "Logging in to Vercel with info@mymagppie.com..."
npx vercel login info@mymagppie.com

# 2. Link using npx vercel
echo "Linking project to Vercel under scope info-38418418s-projects..."
npx vercel link --scope info-38418418s-projects --yes

# 3. Deploy using npx vercel
echo "Deploying project to Vercel..."
npx vercel --prod --yes

echo "=== Deployment Complete! ==="
