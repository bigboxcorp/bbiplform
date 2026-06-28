#!/bin/bash
echo "Starting deployment..."

# Navigate to the project directory
cd ~/bbiplform

echo "Pulling latest code from GitHub..."
git pull origin main

echo "Installing dependencies..."
npm install

echo "Building application..."
npm run build

echo "Restarting PM2 process..."
pm2 restart bbiplform

echo "Saving PM2 process list..."
pm2 save

echo "Deployment complete! Application should be live."
