/**
 * -------------------------------------------------------------------
 * CORE CONTROLLER
 * Handles essential, app-wide data and actions.
 * -------------------------------------------------------------------
 */

/**
 * Fetches the current user's basic information.
 * This is a simple example; in a real app, you might fetch roles or other data.
 * @returns {object} An object containing the user's email and a photo URL.
 */
function api_getUserInfo() {
    try {
        // In a real application, you might also query a 'Staff' sheet to get a custom photo URL or role.
        // For now, we'll use the user's Google account photo if available.
        const email = Session.getActiveUser().getEmail();
        const photoUrl = Session.getActiveUser().getPhotoUrl(); // This might not always be available depending on domain settings

        return {
            success: true,
            data: {
                email: email,
                photoUrl: photoUrl
            }
        };
    } catch (e) {
        console.error("api_getUserInfo Error: " + e.stack);
        // Return a response that allows the UI to still function
        return {
            success: false,
            message: "Could not retrieve user information.",
            data: {
                email: "Error loading user",
                photoUrl: ""
            }
        };
    }
}
