import streamlit as st
import streamlit.components.v1 as components
import os

def render_paypal_button(client_id, amount="99.00", item_name="SCD Dbase Sorter - Lifetime License"):
    """
    Renders the PayPal Smart Button using the JavaScript SDK.
    """
    if not client_id:
        st.warning("⚠️ PayPal Client ID not configured. Billing system inactive.")
        return

    paypal_html = f"""
    <div id="paypal-button-container"></div>
    <script src="https://www.paypal.com/sdk/js?client-id={client_id}&currency=USD"></script>
    <script>
        paypal.Buttons({{
            createOrder: function(data, actions) {{
                return actions.order.create({{
                    purchase_units: [{{
                        amount: {{
                            value: '{amount}'
                        }},
                        description: '{item_name}'
                    }}]
                }});
            }},
            onApprove: function(data, actions) {{
                return actions.order.capture().then(function(details) {{
                    alert('Transaction completed by ' + details.payer.name.given_name + '!');
                    // In a real app, you would notify the Streamlit backend here
                    window.parent.postMessage({{
                        type: 'streamlit:setComponentValue',
                        value: 'SUCCESS'
                    }}, '*');
                }});
            }},
            onError: function(err) {{
                console.error('PayPal Error:', err);
            }}
        }}).render('#paypal-button-container');
    </script>
    """
    
    st.markdown(f"### Upgrade to Pro")
    st.write(f"Get the **{item_name}** for a one-time payment of **${amount}**.")
    components.html(paypal_html, height=350)
