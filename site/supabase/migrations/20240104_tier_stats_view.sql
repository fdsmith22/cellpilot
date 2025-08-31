-- Create a view for tier statistics (optional optimization for large user bases)
CREATE OR REPLACE VIEW tier_statistics AS
SELECT 
  subscription_tier,
  COUNT(*) as user_count,
  AVG(operations_used) as avg_operations_used,
  SUM(operations_used) as total_operations_used
FROM profiles
GROUP BY subscription_tier;

-- Create a function to get all tier counts at once
CREATE OR REPLACE FUNCTION get_tier_counts()
RETURNS TABLE (
  free_count bigint,
  beta_count bigint,
  pro_count bigint,
  enterprise_count bigint,
  total_count bigint
) AS $$
BEGIN
  RETURN QUERY
  SELECT 
    COUNT(*) FILTER (WHERE subscription_tier = 'free') as free_count,
    COUNT(*) FILTER (WHERE subscription_tier = 'beta') as beta_count,
    COUNT(*) FILTER (WHERE subscription_tier = 'pro') as pro_count,
    COUNT(*) FILTER (WHERE subscription_tier = 'enterprise') as enterprise_count,
    COUNT(*) as total_count
  FROM profiles;
END;
$$ LANGUAGE plpgsql;

-- Usage: SELECT * FROM get_tier_counts();