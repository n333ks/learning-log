"""
Change order procedure data and scenario resolution logic.
Source: change_orders/change_order_procedures.json
"""

SCENARIOS = {
    '1A': {
        'title': 'Order is still in the warehouse',
        'part': 'A',
        'part_title': 'Stock / Presale Orders',
        'status_label': 'Change can be made',
        'friction_level': 'low',
        'why_it_exists': (
            "If an order hasn't left the building, a change is the least disruptive outcome "
            "for everyone. No shipping costs are wasted and the customer gets what they want."
        ),
        'cs_steps': [
            "Direct the customer to reply to their order confirmation email with all requested "
            "changes in writing. If by phone, still require email confirmation before acting.",
            "Notify the warehouse manager that the order has been flagged for a change and request "
            "that all actions taken be reversed.",
            "Use the edit function in Shopify to update the order per the customer's request.",
            "If a balance is owed, issue a balance invoice and hold until payment is received.",
            "Update the order card in PTv2 to reflect all changes and return the updated card "
            "to the warehouse.",
        ],
        'warehouse_actions': [
            "Stop all activity on the order immediately when CS flags a change — do not pick, "
            "pack, or stage further.",
            "Reverse any actions already taken on the order (re-shelve pulled items, un-stage, etc.).",
            "Do not process the order again until the updated PTv2 card is received from CS.",
        ],
        'owner_verify': [
            "Customer email reply on file before any action was taken.",
            "PTv2 card update timestamp follows Shopify edit.",
        ],
    },
    '1B': {
        'title': 'Order already shipped via LTL',
        'part': 'A',
        'part_title': 'Stock / Presale Orders',
        'status_label': 'Reversal possible — fees apply',
        'friction_level': 'medium',
        'why_it_exists': (
            "LTL carriers can reverse a shipment mid-transit but it costs money. The customer is "
            "responsible for all fees because the order was correct when it left."
        ),
        'cs_steps': [
            "Inform the customer that the order has already shipped.",
            "Advise the customer that reverse shipping fees back to PINKYS are required on top of "
            "original shipping already paid.",
            "Upon receipt of reverse shipping payment, coordinate the return with the LTL carrier.",
            "Once the order is back at PINKYS, adjust the order and advise the customer an additional "
            "outbound shipping fee applies for the new shipment.",
        ],
        'warehouse_actions': [
            "No floor action needed until the returned shipment arrives back at PINKYS.",
            "When it arrives, treat as standard inbound — inspect, receive, and notify CS.",
            "CS coordinates with the LTL carrier — warehouse role begins when product is back "
            "in the building.",
        ],
        'owner_verify': [
            "Reverse shipping payment confirmed before carrier return was arranged.",
            "Additional outbound shipping fee collected before new shipment released.",
        ],
    },
    '1C': {
        'title': 'Order already shipped via non-LTL',
        'part': 'A',
        'part_title': 'Stock / Presale Orders',
        'status_label': 'No reversal possible — 25% restocking fee',
        'friction_level': 'high',
        'warning': (
            "Non-LTL carriers do not offer mid-transit reversals. The order must complete "
            "delivery before anything can happen."
        ),
        'why_it_exists': (
            "Non-LTL shipments cannot be recalled once they leave the facility. The 25% restocking "
            "fee covers the labor and handling cost of receiving, inspecting, and re-processing "
            "a returned item."
        ),
        'cs_steps': [
            "Inform the customer that the shipment cannot be reversed and delivery must be accepted.",
            "Advise the customer that a 25% restocking fee applies and return shipping to PINKYS "
            "is their responsibility.",
            "Issue an RMA to the customer with all required return instructions.",
            "Upon return, inspect for damage. If no damage found, process a refund less the "
            "25% restocking fee.",
        ],
        'warehouse_actions': [
            "When the return arrives with an RMA, treat as priority inbound — inspect thoroughly "
            "for shipping damage.",
            "Document condition accurately. If damage is found, notify CS immediately before any "
            "refund is issued.",
            "Do not accept a return without a valid RMA number. If no RMA, direct to CS before "
            "touching the product.",
        ],
        'owner_verify': [
            "RMA on file for the return.",
            "Inspection record exists and predates the refund.",
            "Refund amount reflects 25% restocking fee deduction.",
        ],
    },
    '1D': {
        'title': 'Order already picked up or locally delivered',
        'part': 'A',
        'part_title': 'Stock / Presale Orders',
        'status_label': '25% restocking fee applies',
        'friction_level': 'high',
        'why_it_exists': (
            "Once a customer has physical possession, the return process is identical to a standard "
            "return. Same 25% restocking fee applies for the same reasons as a shipped order."
        ),
        'cs_steps': [
            "Advise the customer that a 25% restocking fee applies and return transportation to "
            "PINKYS is their responsibility.",
            "Issue an RMA with all required return instructions.",
            "Upon return, inspect for damage. If no damage found, process a refund less the "
            "25% restocking fee.",
        ],
        'warehouse_actions': [
            "Same inspection process as a shipped return when the product arrives.",
            "Local pickup returns may arrive without a carrier — customer may bring it back "
            "themselves. Still require RMA paperwork before accepting.",
            "Do not accept a return without a valid RMA number. Direct to CS if none present.",
        ],
        'owner_verify': [
            "RMA on file for the return.",
            "Inspection record exists and predates the refund.",
            "Refund amount reflects 25% restocking fee deduction.",
        ],
    },
    'B1C': {
        'title': 'Production has not yet started',
        'part': 'B',
        'part_title': 'Custom / Special Orders',
        'status_label': 'No-fee change possible',
        'friction_level': 'low',
        'why_it_exists': (
            "If production hasn't started, a change is essentially free — no materials cut, no "
            "labor spent. The window for a no-cost change is narrow and closes fast."
        ),
        'cs_steps': [
            "Confirm with the production team that no fees apply and the change can be accommodated.",
            "Communicate to the customer that the change will be made at no charge.",
            "Obtain the customer's request in writing via their order confirmation email.",
            "Update Shopify, PTv3, and all CAD/order files. Notify Nikos and all relevant team members.",
            "Take a screenshot once all systems reflect the completed changes.",
        ],
        'warehouse_actions': [
            "If a custom order is staged or queued, CS may contact you to confirm it hasn't "
            "entered production prep.",
            "Accurate status reporting is critical — incorrect status reporting can lead to costly "
            "production mistakes.",
        ],
        'owner_verify': [
            "Production team confirmation on file before customer was notified.",
            "PTv3 screenshot exists confirming all system updates.",
        ],
    },
    'B1A': {
        'title': 'Production can accommodate the change',
        'part': 'B',
        'part_title': 'Custom / Special Orders',
        'status_label': 'Change possible — fees may apply',
        'friction_level': 'medium',
        'why_it_exists': (
            "When production has started but not gone too far, a change may still be possible but "
            "costs something. Change fees cover re-work costs including materials, CAD re-draws, "
            "and labor re-allocation."
        ),
        'cs_steps': [
            "Contact the production team to confirm feasibility and whether change fees apply.",
            "Communicate the outcome to the customer including any fees and lead time impact.",
            "Obtain written confirmation from the customer via reply to their order confirmation email.",
            "If fees apply, issue an invoice and hold until payment is received before notifying "
            "production to proceed.",
            "Update Shopify and PTv3 to reflect all changes. Take a screenshot confirming updates "
            "are reflected in the system.",
        ],
        'warehouse_actions': [
            "Hold all activity on a flagged custom order until updated CAD files and a confirmed "
            "green light from CS are received.",
            "Do not proceed on verbal direction alone — updated PTv3 documentation must accompany "
            "any change.",
        ],
        'owner_verify': [
            "Change fee invoice and payment receipt on file.",
            "Production notified only after payment was confirmed.",
            "PTv3 screenshot on file.",
        ],
    },
    'B1B': {
        'title': 'Production cannot accommodate the change',
        'part': 'B',
        'part_title': 'Custom / Special Orders',
        'status_label': 'Change not possible — customer decides',
        'friction_level': 'high',
        'warning': (
            "Do not make any changes if production is too far along. Present options to the "
            "customer only. No floor action until customer decision is documented."
        ),
        'why_it_exists': (
            "At a certain production stage, making a change would mean scrapping significant work "
            "and materials. The customer chooses to keep the order or accept whatever limited "
            "modification is still possible."
        ),
        'cs_steps': [
            "Inform the customer that the requested change cannot be made due to production status.",
            "Present options and request a decision in writing: Option A — retain the order as "
            "originally placed (no fee). Option B — accept any partial modification still available, "
            "subject to applicable fees.",
            "Document the customer's decision and update the order record accordingly.",
        ],
        'warehouse_actions': [
            "Continue production as normal unless CS explicitly instructs a stop.",
            "If unsure whether a hold has been placed, check with manager before proceeding — "
            "do not guess.",
        ],
        'owner_verify': [
            "Customer decision documented in writing before any floor action was taken.",
            "No unauthorized changes made to production order.",
        ],
    },
}

NON_NEGOTIABLES = [
    "No change released before payment confirmed — PTv2/PTv3 update timestamp must always follow "
    "payment confirmation, never precede it.",
    "No refund issued before inspection completed — if damage is found after a refund is already "
    "issued, PINKYS absorbs that cost.",
    "No custom order change promised before production confirms feasibility — CS must contact "
    "production before communicating anything to the customer.",
    "All fees are paid by the customer — restocking, reverse shipping, and change fees are always "
    "customer-borne.",
]

FRICTION_COLOR = {
    'low':    ('#E8F5E9', '#2E7D32'),
    'medium': ('#FFF8E1', '#F57F17'),
    'high':   ('#FFEBEE', '#C62828'),
}


def get_scenario(scenario_id):
    return SCENARIOS.get(scenario_id)


def determine_scenario(order_type, stock_location=None,
                       production_started=None, can_accommodate=None):
    if order_type == 'custom':
        if production_started == 'no':
            return 'B1C'
        if can_accommodate == 'yes':
            return 'B1A'
        return 'B1B'
    return stock_location  # '1A' / '1B' / '1C' / '1D'
