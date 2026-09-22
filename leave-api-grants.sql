-- 请在 CloudBase PostgreSQL SQL 编辑器执行一次。
-- HTTP 云函数使用服务端 API Key 访问这两张表。
ALTER TABLE public.leave_forms ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.leave_requests ENABLE ROW LEVEL SECURITY;

GRANT SELECT, INSERT, UPDATE, DELETE ON TABLE public.leave_forms TO service_role;
GRANT SELECT, INSERT, UPDATE, DELETE ON TABLE public.leave_requests TO service_role;
GRANT SELECT, INSERT, UPDATE ON TABLE public.leave_forms TO authenticated;
GRANT SELECT, UPDATE ON TABLE public.leave_requests TO authenticated;

DROP POLICY IF EXISTS leave_forms_select_own ON public.leave_forms;
DROP POLICY IF EXISTS leave_forms_insert_own ON public.leave_forms;
DROP POLICY IF EXISTS leave_forms_update_own ON public.leave_forms;
DROP POLICY IF EXISTS leave_requests_select_own ON public.leave_requests;
DROP POLICY IF EXISTS leave_requests_update_own ON public.leave_requests;

CREATE POLICY leave_forms_select_own ON public.leave_forms
FOR SELECT TO authenticated USING (owner_id = auth.uid());
CREATE POLICY leave_forms_insert_own ON public.leave_forms
FOR INSERT TO authenticated WITH CHECK (owner_id = auth.uid());
CREATE POLICY leave_forms_update_own ON public.leave_forms
FOR UPDATE TO authenticated USING (owner_id = auth.uid()) WITH CHECK (owner_id = auth.uid());
CREATE POLICY leave_requests_select_own ON public.leave_requests
FOR SELECT TO authenticated USING (owner_id = auth.uid());
CREATE POLICY leave_requests_update_own ON public.leave_requests
FOR UPDATE TO authenticated USING (owner_id = auth.uid()) WITH CHECK (owner_id = auth.uid());
