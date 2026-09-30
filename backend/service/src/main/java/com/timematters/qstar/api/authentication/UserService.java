package com.timematters.qstar.api.authentication;

import com.timematters.error.InternalServerError;
import com.timematters.qstar.api.authentication.user.AuthenticatedUser;
import com.timematters.qstar.api.authentication.user.UserId;
import java.util.Optional;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;

/** Unwired legacy template example. No application identity is used by Q-Star APIs. */
public class UserService {
    private UserRepository userRepository;

    private static final Logger log = LoggerFactory.getLogger(UserService.class);

    @Autowired
    public UserService(UserRepository userRepository) {
        this.userRepository = userRepository;
    }

    public UserId getCurrentUserId() {
        return this.userRepository.currentUserId();
    }

    public AuthenticatedUser getCurrentUser() throws InternalServerError {
        try {
            return this.userRepository.getCurrentUser();
        } catch (Exception e) {
            log.error(
                    "Fetching current user failed. "
                            + "This means that something serious went wrong. A authenticated User"
                            + "should always be fetchable from the Active Directory. {}",
                    e);
            throw new InternalServerError(
                    "Fetching current user failed. "
                            + "This means that something serious went wrong. A authenticated User"
                            + "should always be fetchable from the Active Directory.");
        }
    }

    public Optional<AuthenticatedUser> getUserInfo(UserId userId) {
        try {
            return this.userRepository.getUserInfo(userId);

        } catch (Exception e) {
            log.error(
                    "Attempted and failed to fetch infos for userId:"
                            + userId.toString()
                            + " Please investigate why the call failed.",
                    e);
        }
        return Optional.empty();
    }

    public Optional<AuthenticatedUser> getUserInfo(String userId) {
        try {
            return this.getUserInfo(UserId.fromString(userId));
        } catch (IllegalArgumentException e) {
            log.info("{} isn't a valid UUID so no request against the AD will be done.", userId);
        }
        return Optional.empty();
    }
}
