import passport from "passport";
import { Strategy as GoogleStrategy } from "passport-google-oauth20";

const allowedDomain = (process.env.GOOGLE_ALLOWED_DOMAIN || "").toLowerCase().trim();

passport.serializeUser((user, done) => done(null, user));
passport.deserializeUser((user, done) => done(null, user));

passport.use(
  new GoogleStrategy(
    {
      clientID: process.env.GOOGLE_CLIENT_ID,
      clientSecret: process.env.GOOGLE_CLIENT_SECRET,
      callbackURL: `${process.env.BASE_URL}/auth/google/callback`,
    },
    async (_accessToken, _refreshToken, profile, done) => {
      try {
        const email = profile?.emails?.[0]?.value?.toLowerCase() || "";
        const domain = email.split("@")[1] || "";

        // ✅ Domain restriction (real)
        if (!email || domain !== allowedDomain) return done(null, false);

        return done(null, {
          email,
          name: profile.displayName || email,
          photo: profile?.photos?.[0]?.value || null,
          provider: "google",
        });
      } catch (e) {
        return done(e);
      }
    }
  )
);

export default passport;