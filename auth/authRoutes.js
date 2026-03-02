import express from "express";
import passport from "../auth/passport.js";

const router = express.Router();

router.get(
  "/google",
  passport.authenticate("google", {
    scope: ["profile", "email"],
    prompt: "select_account",
    hd: process.env.GOOGLE_ALLOWED_DOMAIN, // hint only
  })
);

router.get(
  "/google/callback",
  passport.authenticate("google", {
    failureRedirect: "/login?err=domain",
    successRedirect: "/",
  })
);

router.post("/logout", (req, res) => {
  req.logout(() => {
    req.session?.destroy(() => res.redirect("/login"));
  });
});

export default router;