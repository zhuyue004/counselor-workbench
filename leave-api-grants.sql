-- 请在 CloudBase PostgreSQL SQL 编辑器执行一次。
-- HTTP 云函数使用服务端 API Key 访问这两张表。
ALTER TABLE public.leave_forms ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.leave_requests ENABLE ROW LEVEL SECURITY;

GRANT SELECT, INSERT, UPDATE, DELETE ON TABLE public.leave_forms TO service_role;
GRANT SELECT, INSERT, UPDATE, DELETE ON TABLE public.leave_requests TO service_role;
