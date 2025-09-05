# Supabase CLI Connection Guide

## Prerequisites
- Supabase CLI installed (`npm install -g supabase` or use `npx supabase`)
- Database password from Supabase Dashboard
- Access token from https://supabase.com/dashboard/account/tokens

## Connection Details
- **Project Ref**: `grhcxavwagumzceauthr`
- **Direct Database Host**: `db.grhcxavwagumzceauthr.supabase.co`
- **Pooler Host**: `aws-1-eu-west-2.pooler.supabase.com`
- **Database Name**: `postgres`
- **Database User**: `postgres`
- **Database Password**: Stored in `.env.local` as `POSTGRES_PASSWORD`

## Initial Setup

1. **Login to Supabase CLI**:
```bash
npx supabase login --token YOUR_ACCESS_TOKEN
```

2. **Link Project** (from project root):
```bash
npx supabase link --project-ref grhcxavwagumzceauthr
# Enter database password when prompted
```

## Database Commands

### Using psql directly:
```bash
# Set password and connect
PGPASSWORD=YOUR_PASSWORD psql -h db.grhcxavwagumzceauthr.supabase.co -U postgres -d postgres

# Query examples
PGPASSWORD=YOUR_PASSWORD psql -h db.grhcxavwagumzceauthr.supabase.co -U postgres -d postgres -c "SELECT * FROM profiles;"
```

### Using Supabase CLI:
```bash
# Database statistics
npx supabase inspect db db-stats --db-url "postgresql://postgres:PASSWORD@db.grhcxavwagumzceauthr.supabase.co:5432/postgres"

# Table statistics
npx supabase inspect db table-stats --db-url "postgresql://postgres:PASSWORD@db.grhcxavwagumzceauthr.supabase.co:5432/postgres"

# Test connection
npx supabase test db --db-url "postgresql://postgres:PASSWORD@db.grhcxavwagumzceauthr.supabase.co:5432/postgres"
```

## Important Notes
- Use direct database host (`db.grhcxavwagumzceauthr.supabase.co`) for admin operations
- Use pooler host for application connections
- Port 5432 for direct database
- Port 6543 for pooler connections
- Always use SSL in production

## Common Issues
- If connection refused on pooler, use direct database host
- If password authentication fails, reset in Supabase Dashboard
- Docker image downloads may timeout - be patient or use psql directly