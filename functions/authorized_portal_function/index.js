const catalyst = require("zcatalyst-sdk-node");

function writeUserResponse(basicIO, context, userDetails) {
  const role = userDetails.role_identifier || userDetails.role || "App Administrator";
  const emailId = userDetails.email_id || userDetails.email || "";
  const output = {
    status: "success",
    email_id: emailId,
    role,
    user_details: {
      first_name: userDetails.first_name || "",
      last_name: userDetails.last_name || "",
      email_id: emailId,
      role_identifier: role,
      role,
      user_id: userDetails.user_id || userDetails.id || "",
      account_id: userDetails.account_id || userDetails.zaaid || "",
      org_id: userDetails.org_id || "",
      created_time: userDetails.created_time || userDetails.created_time_long || "",
      status: userDetails.status || "",
    },
  };
  basicIO.write(JSON.stringify(output));
  context.close();
}

module.exports = (context, basicIO) => {
  const catalystApp = catalyst.initialize(context);
  const userManagement = catalystApp.userManagement();

  // When called via HTTP GET (from the app), get the current logged-in user – return profile for any authenticated user so UI can display it
  userManagement.getCurrentUser().then((currentUser) => {
    if (currentUser && (currentUser.email_id || currentUser.email)) {
      const email = (currentUser.email_id || currentUser.email || "").trim();
      const roleName = (currentUser.role_details && currentUser.role_details.name) || currentUser.role_identifier || currentUser.role || "App Administrator";
      writeUserResponse(basicIO, context, {
        first_name: currentUser.first_name || "",
        last_name: currentUser.last_name || "",
        email_id: email,
        role_identifier: roleName,
        role: roleName,
        user_id: currentUser.user_id || currentUser.id || "",
        account_id: currentUser.account_id || currentUser.zaaid || "",
        org_id: currentUser.org_id || "",
        created_time: currentUser.created_time || currentUser.created_time_long || "",
        status: currentUser.status || "",
      });
    } else {
      // No user from getCurrentUser: try signup validation (e.g. during login flow)
      const requestDetails = userManagement.getSignupValidationRequest(basicIO);
      if (requestDetails && requestDetails.user_details) {
        const ud = requestDetails.user_details;
        const email = (ud.email_id || ud.email || "").trim();
        if (email && (email.includes("@buildhr.co.in") || email.includes("@zylker.com"))) {
          writeUserResponse(basicIO, context, {
            first_name: ud.first_name || "",
            last_name: ud.last_name || "",
            email_id: email,
            role_identifier: ud.role_identifier || "App Administrator",
            role: ud.role_identifier || ud.role || "App Administrator",
          });
        } else {
          basicIO.write(JSON.stringify({ status: "failure" }));
          context.close();
        }
      } else {
        basicIO.write(JSON.stringify({ status: "failure", message: "Not authenticated" }));
        context.close();
      }
    }
  }).catch((err) => {
    basicIO.write(JSON.stringify({ status: "failure", message: "Error fetching user", error: String(err && err.message || err) }));
    context.close();
  });
};
