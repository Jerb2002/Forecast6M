# What the Check sheet does under an override

Type: grilling
Status: open
Blocked by: 05

## Question

The user has already set the rule: the checks keep tying CF to Budget on the **budget timeline
only**, exactly as they do now. A month under an override is not expected to add up.

What is not settled is what the Check sheet *shows* when the dropdown is not on Budget only.

- Does it pin itself to the budget basis regardless of the switch, so it always reconciles the one
  thing it can reconcile - and then risks reading as a green tick over a sheet nobody is looking at?
- Or does it gain a second block that states the override gap plainly, so the un-tied amount is
  reported rather than hidden?

The second is close to what the user wants out of *Difference* on the dropdown, so settle whether
these are one feature or two before building either.

Also: the three VAT rows must still sum to nothing over the horizon - what we collect, less what we
pay, is what we settle. An override drops money onto the revenue row without a matching invoice
behind it. Work out whether that identity survives, and if it does not, whether that is a check
which should fire or a check which should exclude overridden weeks.
