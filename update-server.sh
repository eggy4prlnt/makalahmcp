#!/bin/bash

# Script untuk update MakalahMCP di server
# Jalankan di server: ./update-server.sh

set -e

echo "🔄 Updating MakalahMCP Server"
echo "=============================="
echo ""

# Check if we're in the right directory
if [ ! -f "main.py" ]; then
    echo "❌ Error: main.py not found"
    echo "   Please run this script from /root/makalahmcp directory"
    exit 1
fi

echo "📍 Current directory: $(pwd)"
echo ""

# Stop running containers
echo "1. Stopping running containers..."
if docker-compose ps | grep -q "Up"; then
    docker-compose down
    echo "   ✅ Containers stopped"
else
    echo "   ℹ️  No running containers"
fi
echo ""

# Pull latest changes and remove deleted files
echo "2. Pulling latest changes from git..."
git fetch origin

# Show what will be deleted
echo ""
echo "   Files that will be removed:"
git diff --name-status HEAD origin/main | grep "^D" | awk '{print "   - " $2}' || echo "   (none)"
echo ""

# Reset to match remote exactly (removes deleted files)
git reset --hard origin/main

# Clean untracked files
git clean -fd

echo "   ✅ Code updated to latest version"
echo "   ✅ Deleted files removed"
echo ""

# Show current commit
CURRENT_COMMIT=$(git log -1 --oneline)
echo "   Current commit: $CURRENT_COMMIT"
echo ""

# Rebuild containers
echo "3. Rebuilding Docker containers..."
docker-compose build --no-cache
echo "   ✅ Containers rebuilt"
echo ""

# Start containers
echo "4. Starting containers..."
docker-compose up -d
echo "   ✅ Containers started"
echo ""

# Wait for services to be ready
echo "5. Waiting for services to be ready..."
sleep 5

# Check health
echo ""
echo "6. Checking service health..."
if curl -s http://localhost:8000/health > /dev/null 2>&1; then
    echo "   ✅ MakalahMCP server is healthy"
else
    echo "   ⚠️  MakalahMCP server is not responding yet"
fi
echo ""

# Show logs
echo "=============================="
echo "✅ Update Complete!"
echo "=============================="
echo ""
echo "📊 Service Status:"
docker-compose ps
echo ""
echo "📝 View logs:"
echo "   docker-compose logs -f"
echo ""
echo "🌐 Service URL:"
echo "   http://localhost:8000"
echo ""
