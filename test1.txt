import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.web.bind.annotation.*;
import org.springframework.beans.factory.annotation.Autowired;
 
import javax.persistence.*;
 
@SpringBootApplication
public class FraudDataApiApplication {
    public static void main(String[] args) {
        SpringApplication.run(FraudDataApiApplication.class, args);
    }
}
 
@Entity // This tells Hibernate to make a table out of this class
@Table(name = "fraud_data") // Name of the table in the database
class FraudData {
 
    @Id // This tells Hibernate to make this field the primary key
    @GeneratedValue(strategy=GenerationType.AUTO) // This tells Hibernate to auto-increment the primary key
    private Integer id;
 
    private String type;
 
    private String description;
 
    // getters and setters here
 
}
 
@RestController // This tells Spring Boot that this class is a REST controller
@RequestMapping("/fraud-data") // This tells Spring Boot that all requests to /fraud-data will be handled by this controller
class FraudDataController {
 
    @Autowired // This tells Spring Boot to inject the FraudDataRepository into this controller
    private FraudDataRepository fraudDataRepository;
 
    @PostMapping // This tells Spring Boot that this method will handle POST requests to /fraud-data
    public FraudData createFraudData(@RequestBody FraudData fraudData) {
        return fraudDataRepository.save(fraudData);
    }
 
    @GetMapping // This tells Spring Boot that this method will handle GET requests to /fraud-data
    public List<FraudData> getAllFraudData() {
        return fraudDataRepository.findAll();
    }
 
    @GetMapping("/{id}") // This tells Spring Boot that this method will handle GET requests to /fraud-data/{id}
    public FraudData getFraudDataById(@PathVariable Integer id) {
        return fraudDataRepository.findById(id).orElseThrow(() -> new ResourceNotFoundException("Fraud data not found with id :" + id));
    }
 
    @PutMapping("/{id}") // This tells Spring Boot that this method will handle PUT requests to /fraud-data/{id}
    public FraudData updateFraudData(@PathVariable Integer id, @RequestBody FraudData fraudData) {
        FraudData fraudDataToUpdate = fraudDataRepository.findById(id).orElseThrow(() -> new ResourceNotFoundException("Fraud data not found with id :" + id));
        fraudDataToUpdate.setType(fraudData.getType());
        fraudDataToUpdate.setDescription(fraudData.getDescription());
        return fraudDataRepository.save(fraudDataToUpdate);
    }
 
    @DeleteMapping("/{id}") // This tells Spring Boot that this method will handle DELETE requests to /fraud-data/{id}
    public void deleteFraudData(@PathVariable Integer id) {
        fraudDataRepository.deleteById(id);
    }
 
}
 
@Repository // This tells Spring Boot that this class is a JPA repository
interface FraudDataRepository extends JpaRepository<FraudData, Integer> {
 
}
