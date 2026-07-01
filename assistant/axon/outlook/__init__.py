"""Outlook automation package. Re-exports the tool functions + helpers other modules import."""
from axon.outlook._base import _run_outlook_ps
from axon.outlook.identity import _self_email, _RESOLVE_EMAIL_PS, _SELF_EMAIL_PS
from axon.outlook.email_send import (
    _send_email, _send_draft, _draft_then_send, _outlook_reply_email, _outlook_forward_email,
)
from axon.outlook.messages import (
    _outlook_list_emails, _outlook_move_emails, _outlook_delete_emails, _show_email,
    _outlook_save_email, _get_open_email, _outlook_mark_read, _outlook_categorize,
)
from axon.outlook.folders import (
    _create_outlook_folders, _delete_outlook_folders, _list_outlook_folders, _create_outlook_rule,
)
from axon.outlook.calendar import _create_calendar_event
from axon.outlook.signature import _set_outlook_signature
from axon.outlook.tone import learn_my_tone, my_tone
