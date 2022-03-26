package models;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.math.BigDecimal;
import java.time.LocalDateTime;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class User {
    private Long id;
    private String name;
    private Role role;
    private LocalDateTime registerDate;
    private boolean active;
    private BigDecimal balance;
    
    public enum Role {
        USER, ADMIN
    }
}
